using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
 
/// <summary>
/// HTMLParser is an object that can decode mhtml into ASCII text.
/// Using getHTMLText() will generate static HTML with inline images. 
/// </summary>
public class MHTMLParser
{
    const string BOUNDARY = "boundary";
    const string CHAR_SET = "charset";
    const string CONTENT_TYPE = "Content-Type";
    const string CONTENT_TRANSFER_ENCODING = "Content-Transfer-Encoding";
    const string CONTENT_LOCATION = "Content-Location";
    const string FILE_NAME = "filename=";
 
    private string mhtmlString; // the string we want to decode
    private string log; // log file
    public bool decodeImageData; //decode images?
 
    /*
     * Results of Conversion
     * This is split into a string[3] for each part
     * string[0] is the content type
     * string[1] is the content name
     * string[2] is the converted data
     */
    public List<string[]> dataset;
 
    /*
     * Default Constructor
     */
    public MHTMLParser()
    {
        dataset = new List<string[]>(); //Init dataset
        log += "Initialized dataset.\n";
        decodeImageData = false; //Set default for decoding images
    }
 
    /*
     * Init with contents of string 
     */
    public MHTMLParser(string mhtml)
        : this()
    {
        SetMHTMLString(mhtml);
    }
    /*
     * Init with contents of string, and decoding option
     */
    public MHTMLParser(string mhtml, bool decodeImages)
        : this(mhtml)
    {
        decodeImageData = decodeImages;
    }
    /*
     * Set the mhtml string we want to decode
     */
    public void SetMHTMLString(string mhtml)
    {
        try
        {
            mhtmlString = mhtml ?? throw new Exception("The mhtml string is null"); //Set String
            log += "Set mhtml string.\n";
        }
        catch (Exception e)
        {
            log += e.Message;
            log += e.StackTrace;
        }
    }
    /*
     * Decompress Archive From String
     */
    public List<string[]> DecompressString()
    {
        // init Prerequisites
        StringReader reader = null;
        string type = "";
        string encoding = "";
        string location = "";
        string filename = "";
        string charset = "utf-8";
        StringBuilder buffer = null;
        log += "Starting decompression \n";
 
 
        try
        {
            reader = new StringReader(mhtmlString); //Start reading the string
 
            string boundary = GetBoundary(reader); // Get the boundary code
            if (boundary == null) throw new Exception("Failed to find string 'boundary'");
            log += "Found boundary.\n";
 
            //Loop through each line in the string
            string line = null;
            while ((line = reader.ReadLine()) != null)
            {
                string temp = line.Trim();
                if (temp.Contains(boundary)) //Check if this is a new section
                {
                    if (buffer != null) //If this is a new section and the buffer is full, write to dataset
                    {
                        string[] data = new string[3];
                        data[0] = type;
                        data[1] = filename!=""?filename:location;
                        data[2] = WriteBufferContent(buffer, encoding, charset, type, decodeImageData);
                        dataset.Add(data);
                        buffer = null;
                        log += "Wrote Buffer Content and reset buffer.\n";
                    }
                    buffer = new StringBuilder();
                }
                else if (temp.StartsWith(CONTENT_TYPE))
                {
                    type = GetAttribute(temp);
                    log += "Got content type.\n";
                }
                else if (temp.StartsWith(CHAR_SET))
                {
                    charset = GetCharSet(temp);
                    log += "Got charset.\n";
                }
                else if (temp.StartsWith(CONTENT_TRANSFER_ENCODING))
                {
                    encoding = GetAttribute(temp);
                    log += "Got encoding (" + encoding + ").\n";
                }
                else if (temp.StartsWith(CONTENT_LOCATION))
                {
                    location = temp.Substring(temp.IndexOf(":") + 1).Trim();
                    log += "Got location.\n";
                }
                else if (temp.StartsWith(FILE_NAME))
                {
                    char c = '"';
                    filename = temp.Substring(temp.IndexOf(c.ToString()) + 1, temp.LastIndexOf(c.ToString()) - temp.IndexOf(c.ToString()) - 1);
                }
                else if (temp.StartsWith("Content-ID") || temp.StartsWith("Content-Disposition") || temp.StartsWith("name=") || temp.Length == 1)
                {
                    //We don't need this stuff; Skip lines
                }
                else
                {
                    if (buffer != null)
                    {
                        buffer.Append(line + "\n");
                    }
                }
            }
        }
        finally
        {
            if (null != reader)
                reader.Close();
            log += "Closed Reader.\n";
        }
        return dataset; //Return Results
    }
    private string WriteBufferContent(StringBuilder buffer, string encoding, string charset, string type, bool decodeImages)
    {
        log += "Start writing buffer contents.\n";
 
        //Detect if this is an image and if we want to decode it
        if (type.Contains("image"))
        {
            log += "Image Data Detected.\n";
            if (!decodeImages)
            {
                log += "Skipping image decode.\n";
                return buffer.ToString();
            }
        }
 
        // base64 Decoding
        if (encoding.ToLower().Equals("base64"))
        {
            try
            {
                log += "base64 encoding detected.\n";
                log += "Got base64 decoded string.\n";
                return DecodeFromBase64(buffer.ToString());
            }
            catch (Exception e)
            {
                log += e.Message + "\n";
                log += e.StackTrace + "\n";
                log += "Data not Decoded.\n";
                return buffer.ToString();
            }
        }
        //quoted-printable decoding
        else if (encoding.ToLower().Equals("quoted-printable"))
        {
            log += "Quoted-Prinatble string detected.\n";
            return GetQuotedPrintableString(buffer.ToString());
        }
        else
        {
            log += "Unknown Encoding.\n";
            return buffer.ToString();
        }
    }
    /*
     * Take base64 string, get bytes and convert to ascii string
     */
    public string DecodeFromBase64(string encodedData)
    {
        byte[] encodedDataAsBytes
            = System.Convert.FromBase64String(encodedData);
        string returnValue =
           System.Text.ASCIIEncoding.ASCII.GetString(encodedDataAsBytes);
        return returnValue;
    }
    /*
     * Get decoded quoted printable string
     */
    public string GetQuotedPrintableString(string mimeString)
    {
        try
        {
            throw new Exception("Quoted-Printable is not supported.");
        }
        catch (Exception e)
        {
            log += e.Message + "\n";
            log += e.StackTrace + "\n";
            log += "Data not Decoded.\n";
            return mimeString;
        }
    }
    /*
     * Finds boundary used to break code into multiple parts
     */
    private string GetBoundary(StringReader reader)
    {
        string line = null;
 
        while ((line = reader.ReadLine()) != null)
        {
            line = line.Trim();
            //If the line starts with BOUNDARY, lets grab everything in quotes and return it
            if (line.StartsWith(BOUNDARY))
            {
                char c = '"';
                int a = line.IndexOf(c.ToString());
                int b = line.LastIndexOf(c.ToString());
                return line.Substring(line.IndexOf(c.ToString()) + 1, line.LastIndexOf(c.ToString()) - line.IndexOf(c.ToString()) - 1);
            }
        }
        return null;
    }
    /*
     * Grabs charset from a line 
     */
    private string GetCharSet(string temp)
    {
        string t = temp.Split('=')[1].Trim();
        return t.Substring(1, t.Length - 1);
    }
    /*
     * split a line on ": "
     */
    private string GetAttribute(string line)
    {
        string str = ": ";
        return line.Substring(line.IndexOf(str) + str.Length, line.Length - (line.IndexOf(str) + str.Length)).Replace(";", "");
    }
    /*
     * Get an html page from the mhtml. Embeds images as base64 data
     */
    public string GtHTMLText()
    {
        if (decodeImageData) throw new Exception("Turn off image decoding for valid html output.");
        List<string[]> data = DecompressString();
        string body = "";
        //First, lets write all non-images to mail body
        //Then go back and add images in 
        for (int i = 0; i < 2; i++)
        {
            foreach (string[] strArray in data)
            {
                if (i == 0)
                {
                    if (strArray[0].Equals("text/html"))
                    {
                        body += strArray[2];
                        log += "Writing HTML Text\n";
                    }
                }
                else if (i == 1)
                {
                    if (strArray[0].Contains("image"))
                    {
                        body = body.Replace("cid:" + strArray[1], "data:" + strArray[0] + ";base64," + strArray[2]);
                        log += "Overwriting HTML with image: " + strArray[1] + "\n";
                    }
                }
            }
        }
        return body;
    }
    /*
     *  Get the log from the decoding process
     */
    public string GetLog()
    {
        return log;
    }
}

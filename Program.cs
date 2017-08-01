using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace QQMhtToHtml
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Title = "QQMht To Html";
            string mht, html = string.Empty;
            try
            {
                mht = File.ReadAllText(args[0]);

            }
            catch
            {
                return;
            }
            MHTMLParser parser = new MHTMLParser(mht);
            List<string[]> nodes = null;
            try
            {
               Console.WriteLine("Processing data...");
              nodes= parser.DecompressString();
            }
            catch
            {
                return;
            }
            if (nodes.Count > 0)
            {
                int c = nodes.Count - 1;
                if (nodes.Count > 1)
                {
                    Directory.CreateDirectory("images");
                    Console.WriteLine($"{c} image(s) found.");
                }
                html = nodes[0][2];
              for (int i = 1; i < nodes.Count; i++)
                {
                    string ext = nodes[i][0].Split("/".ToArray())[1],
                        name = nodes[i][1].Split(".".ToArray())[0];
                    if (ext == "jpeg")
                        ext = "jpg";
                    string iFile = $@"images\{name}.{ext}";
                    byte[] bytes = Convert.FromBase64String(nodes[i][2]);
                     File.WriteAllBytes(iFile, bytes);
                    html = html.Replace($"{name}.dat", $@"{iFile}");
                    Console.WriteLine($"Processing image...({i}/{c})");
                }
                string path = Path.GetDirectoryName(args[0]);
                string hFile = Path.GetFileNameWithoutExtension(args[0]);
                hFile = $@"{path}\{hFile}.html";
                File.WriteAllText(hFile, html);
                Console.Write("All done.");
                Console.ReadKey();
            }
        }
    }
}
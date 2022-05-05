using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

namespace InternetProvider
{
    internal static class Addons
    {
        static string dbFileName = "a.docx";
        static string wordFileName = "a.accdb";

        static string path = AppDomain.CurrentDomain.BaseDirectory;
        static string downloadFile = "https://github.com/AlexSlz/InternetProvider/raw/master/a.zip";
        static List<string> temp => downloadFile.Split('/').ToList();
        static string downloadFileName => temp[temp.Count - 1].Split('?')[0];

        public static void TryDownload()
        {
            string temp = Path.Combine(path, downloadFileName);

            try
            {
                if ((!File.Exists(dbFileName) || !File.Exists(wordFileName)) && !File.Exists(downloadFileName))
                {
                    using (WebClient client = new WebClient())
                    {
                        client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                        client.DownloadFile(downloadFile, temp);
                    }
                }

                if (File.Exists(downloadFileName))
                {
                    System.IO.Compression.ZipFile.ExtractToDirectory(temp, path);
                    File.Delete(temp);
                }
            }
            catch
            {
            }

        }

        

    }
}

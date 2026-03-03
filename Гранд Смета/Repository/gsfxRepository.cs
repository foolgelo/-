using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Гранд_Смета.Repository
{
    public class gsfxRepository
    {
        private readonly string _folderPath;
        public gsfxRepository(string folderPath)
        {
            if (!Directory.Exists(folderPath))
                _folderPath = folderPath;
        }
        public IEnumerable<XDocument> LoadXMLFromGSFX() 
        {
            XDocument xDoc = null;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Encoding win1251 = Encoding.GetEncoding("windows-1251");

            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(_folderPath))
                {
                    ZipArchiveEntry entry = archive.GetEntry("Data.xml");
                    if (entry != null)
                    {
                        using (Stream stream = entry.Open())
                        {
                            using (var reader = new StreamReader(stream, win1251))
                            {
                                xDoc = XDocument.Load(stream);
                            }                               
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка при чтении архива {_folderPath}: {ex.Message}");
            }

            if (xDoc != null)
            {
                yield return xDoc;
            }
        }
    }
}

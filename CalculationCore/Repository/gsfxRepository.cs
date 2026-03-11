using CalculationCore.Domain;
using CalculationCore.Parsers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace CalculationCore.Repository
{
    public class gsfxRepository
    {
        private readonly string _folderPath;

        public gsfxRepository(string folderPath)
        {
            _folderPath = folderPath;
        }
        public IEnumerable<(string SourcePath, XDocument XDoc)> LoadXMLFromGSFX(string folderPath)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Encoding win1251 = Encoding.GetEncoding("windows-1251");

            if (File.Exists(folderPath))
            {
                XDocument xDoc = ReadXmlData(folderPath, win1251);
                if (xDoc != null) yield return (folderPath, xDoc);
                yield break;
            }

            if (Directory.Exists(folderPath))
            {
                var files = Directory.GetFiles(folderPath, "*.gsfx").ToList();

                foreach (var file in files)
                {
                    XDocument xDoc = ReadXmlData(file, win1251);
                    if (xDoc != null) yield return (file, xDoc);
                }
            }
        }

        private XDocument ReadXmlData(string filePath, Encoding win1251)
        {
            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(filePath))
                {
                    ZipArchiveEntry entry = archive.GetEntry("Data.xml");
                    if (entry != null)
                    {
                        using (Stream stream = entry.Open())
                        using (var reader = new StreamReader(stream, win1251))
                        {
                            return XDocument.Load(reader);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка при чтении архива {filePath}: {ex.Message}");
            }
            return null;
        }

        public IEnumerable<Position> GetAllMaterials()
        {
            foreach (var (sourcePath, xDoc) in LoadXMLFromGSFX(_folderPath))
            {
                string fileName = Path.GetFileNameWithoutExtension(sourcePath);
                var parser = new XMLParser(fileName, xDoc);
                foreach (var pos in parser.ParsePositions())
                {
                    yield return pos;
                }
            }
        }
    }
}


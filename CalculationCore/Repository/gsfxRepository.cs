using CalculationCore.Domain;
using CalculationCore.Factories;
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
        public IEnumerable<(XDocument XDoc, string SourcePath)> LoadXMLFromGSFX()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Encoding win1251 = Encoding.GetEncoding("windows-1251");

            if (File.Exists(_folderPath))
            {
                XDocument xDoc = ReadXmlData(_folderPath, win1251);
                if (xDoc != null) yield return (xDoc, _folderPath);
                yield break;
            }

            if (Directory.Exists(_folderPath))
            {
                var files = Directory.GetFiles(_folderPath, "*.gsfx").ToList();

                foreach (var file in files)
                {
                    XDocument xDoc = ReadXmlData(file, win1251);
                    if (xDoc != null) yield return (xDoc, file);
                }
            }
        }

        private XDocument ReadXmlData(string archivePath, Encoding win1251)
        {
            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(archivePath))
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
                Debug.WriteLine($"Ошибка при чтении архива {archivePath}: {ex.Message}");
            }
            return null;
        }

        public IEnumerable<Material> GetMaterialsFromGSFX()
        {
            var docs = LoadXMLFromGSFX();
            foreach (var (xDoc, sourcePath) in docs)
            {
                string fileName = Path.GetFileNameWithoutExtension(sourcePath);

                var factory = new MaterialFactory();

                var materials = factory.GetCalcFromXML(xDoc, fileName) ?? Enumerable.Empty<Material>();
                foreach (var material in materials)
                {
                    if (material != null)
                        yield return material;
                }
            }
        }
    }
}

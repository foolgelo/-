// See https://aka.ms/new-console-template for more informationus\
using System.IO.Compression;
using System.Reflection.Metadata;
using System.Text;
using System.Xml;
using System.Xml.Linq;


string zipPath = @"C:\Олег\Документы\Газпром\97445_17_00092_Локальная смета.gsfx";
string data = "Data.xml";

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
Encoding win1251 = Encoding.GetEncoding("windows-1251");

try
{
    using (ZipArchive archive = ZipFile.OpenRead(zipPath))
    {
        ZipArchiveEntry? entry = archive.GetEntry(data);
        if (entry != null)
        {
            using (Stream entryStream = entry.Open())
            {
                using (var reader = new StreamReader(entryStream, win1251))
                {
                    XDocument xmlDoc = XDocument.Load(reader);

                    var allElements = xmlDoc.Descendants();

                    foreach (var element in allElements)
                    {
                        Console.Write($"Элемент: {element.Name}");

                        if (element.HasAttributes)
                        {
                            Console.Write($" [Атрибуты: {string.Join(", ", element.Attributes())}]");
                        }

                        if (!string.IsNullOrEmpty(element.Value.Trim()))
                        {
                            Console.Write($" (Значение: {element.Value})");
                        }

                        Console.WriteLine();
                    }
                }
            }

        }
        else
        {
            Console.WriteLine();
        }
    }
}
catch (FileNotFoundException)
{
    Console.WriteLine("Архив не найден");
}
catch (Exception ex)
{
    Console.WriteLine($"Произошла ошибка: {ex.Message}");
}

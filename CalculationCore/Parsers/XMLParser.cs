using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace CalculationCore.Parsers
{
    public class XMLParser
    {
        //public static IEnumerable<T> Parse<T>(XDocument doc, string archiveName, string tag, Func<XElement, T> mapper)
        //                    where T : BaseEntity
        //{
        //    string locNum = (string)doc.Descendants("Properties").FirstOrDefault()?.Attribute("LocNum") ?? "Н/Д";

        //    return doc.Descendants(tag).Select(el => {
        //        T item = mapper(el); // Специфичные поля для каждого типа
        //        item.SourceFileName = archiveName;
        //        item.LocalNumber = locNum;
        //        item.Name = (string)el.Attribute("Caption") ?? "";
        //        return item;
        //    });
        //}
    }
}

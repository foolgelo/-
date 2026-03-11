using CalculationCore.Domain;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace CalculationCore.Parsers
{
    public class CalculationParser
    {
        public static IEnumerable<Material> Parse(XDocument doc, string fileName)
        {
            string locNum = doc.Descendants("Properties").FirstOrDefault()?.Attribute("LocNum")?.Value ?? "Н/Д";

            foreach (var chapter in doc.Descendants("Chapter"))
            {
                string chapterName = (string)chapter.Attribute("Caption") ?? "Общий раздел";

                foreach (var pos in chapter.Descendants("Position"))
                {
                    var resources = pos.Element("Resources");

                    if (resources != null && resources.HasElements)
                    {
                        foreach (var mat in resources.Descendants("Mat"))
                        {
                            var m = CreateMaterial(mat, chapterName, fileName, locNum);
                            yield return m;
                        }
                    }
                    else
                    {
                        var m = CreateMaterial(pos, chapterName, fileName, locNum);
                        yield return m;
                    }
                }
            }
        }

        private static Material CreateMaterial(XElement el, string chapter, string file, string loc)
        {
            return new Material
            {
                Name = (string)el.Attribute("Caption") ?? "",
                Code = (string)el.Attribute("Code") ?? ""
            };
        }
    }
}

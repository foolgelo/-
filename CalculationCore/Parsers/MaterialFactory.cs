using CalculationCore.Domain;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace CalculationCore.Factories
{
    public class MaterialFactory
    {
        private readonly string _tag = "Mat";
        public static Material CreateFromXElement(XElement el, string fileName, string locNum)
        {
            return new Material
            {
                Name = (string)el.Attribute("Caption") ?? "",
                Code = (string)el.Attribute("Code") ?? "",
                Unit = (string)el.Attribute("Units") ?? "",
                Quantity = (string)el.Attribute("Quantity"),
                BasePrice = (decimal)ParseDouble((string)el.Element("PriceBase")?.Attribute("Value")),
                CurrPrice = (decimal)ParseDouble((string)el.Element("PriceCurr")?.Attribute("Value"))
            };
        }
        public IEnumerable<Material> GetCalcFromXML(XDocument xDoc, string fileName)
        {
            string locNum = (string)xDoc.Descendants("Properties")
                            .FirstOrDefault()?
                            .Attribute("LocNum") ?? "Н/Д";


            foreach (var el in xDoc.Descendants("Mat"))
            {
                yield return MaterialFactory.CreateFromXElement(el, fileName, locNum);
            }
        }
        private static double ParseDouble(string value)
        {
            if (string.IsNullOrEmpty(value)) return 0;
            double.TryParse(value.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var result);
            return result;
        }
    }
}

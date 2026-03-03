using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Гранд_Смета.Domain;

namespace Гранд_Смета.Factories
{
    public class MaterialFactory
    {
        private readonly string _tag = "Mat";
        public static Material CreateFromXElement(XElement el)
        {
            return new Material
            {
                Name = (string)el.Attribute("Caption") ?? "",
                Code = (string)el.Attribute("Code") ?? "",
                Unit = (string)el.Attribute("Units") ?? "",
                Quantity = (decimal)ParseDouble((string)el.Attribute("Quantity"))
            };
        }
        public List<Material> GetCalcFromXML(XDocument xDoc)
        {
            return xDoc.Descendants(_tag)
                        .Select(CreateFromXElement)
                        .ToList();
        }
        private static double ParseDouble(string value)
        {
            if (string.IsNullOrEmpty(value)) return 0;
            double.TryParse(value.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var result);
            return result;
        }
    }
}

using CalculationCore.Domain;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace CalculationCore.Parsers
{
    public class XMLParser
    {
        private readonly string _fileName;
        private readonly string _locNum;
        private readonly IEnumerable<XElement> _posElements;

        public XMLParser(string fileName, XDocument xDoc)
        {
            _fileName = fileName;
            _locNum = xDoc.Descendants("Properties").FirstOrDefault()?.Attribute("LocNum")?.Value ?? "Н/Д";
            _posElements = xDoc.Descendants("Position");
        }

        public IEnumerable<Position> ParsePositions()
        {
            foreach (var posEl in _posElements)
            {
                var currPriceAttr = posEl.Element("PriceCurr")?.Attribute("MT");
                if (currPriceAttr != null)
                {
                    var material = new Material
                    {
                        Name = (string)posEl.Attribute("Caption") ?? "",
                        Code = (string)posEl.Attribute("Code") ?? "",
                        Unit = (string)posEl.Attribute("Units") ?? "",
                        Quantity = (string)posEl.Attribute("Quantity") ?? "",
                        CurrPrice = ParseDecimal((string)currPriceAttr)
                    };
                    yield return new Position
                    {
                        FileName = _fileName,
                        LocNumber = _locNum,
                        //Chapter = (string)posEl.Attribute("Chapter") ?? "Н/Д",
                        //Code = (string)posEl.Attribute("Code"),
                        //Name = (string)posEl.Attribute("Caption"),
                        //Unit = (string)posEl.Attribute("Units"),
                        //Quantity = (string)posEl.Attribute("Quantity") ?? "",
                        Materials = new List<Material> { material }
                    };
                }
                else
                {
                    var position = new Position
                    {
                        FileName = _fileName,
                        LocNumber = _locNum,
                        //Chapter = (string)posEl.Attribute("Chapter") ?? "Н/Д",
                        Code = (string)posEl.Attribute("Code"),
                        Name = (string)posEl.Attribute("Caption"),
                        Unit = (string)posEl.Attribute("Units"),
                        Quantity = (string)posEl.Attribute("Quantity") ?? "",
                        Materials = ParseMaterials(posEl).ToList()
                    };
                    yield return position;
                }
            }
        }

        private IEnumerable<Material> ParseMaterials(XElement posEl)
        {
            var matElements = posEl.Descendants("Mat");
            foreach (var mat in matElements)
            {
                yield return new Material
                {
                    Name = (string)mat.Attribute("Caption") ?? "",
                    Code = (string)mat.Attribute("Code") ?? "",
                    Unit = (string)mat.Attribute("Units") ?? "",
                    Quantity = (string)mat.Attribute("Quantity") ?? "",
                    BasePrice = ParseDecimal((string)mat.Element("PriceBase")?.Attribute("Value")),
                    CurrPrice = ParseDecimal((string)mat.Element("PriceCurr")?.Attribute("Value"))
                };
            }
        }

        private decimal ParseDecimal(string value)
        {
            if (string.IsNullOrEmpty(value)) return 0;
            decimal.TryParse(value.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var result);
            return result;
        }
    }
}
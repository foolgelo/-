using CalculationCore.Domain;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace CalculationCore.Parsers
{
    public class XMLCulcBuilder
    {
        private Position _position = new Position();
        private static XElement _xPos;

        public XMLCulcBuilder(string fileName, string locNum, string chapter, XElement xPos)
        {
            _position.FileName = fileName;
            _position.LocNumber = locNum;
            _position.Chapter = chapter;
            _xPos = xPos;
        }

        public void AddPosition()
        {
            BuildBaseProperties(_xPos, _position);
            _position.Quantity = (string)_xPos.Attribute("Quantity") ?? "";
        }

        private void BuildBaseProperties(XElement el, ICalculation culc) 
        {
            culc.Code = (string)el.Attribute("Code");
            culc.Name = (string)el.Attribute("Caption");
            culc.Unit = (string)el.Attribute("Units");
        }

        public void AddMaterials() 
        {
            var matElements = _xPos?.Descendants("Mat");

            if (matElements != null)
            {
                foreach (var mat in matElements)
                {
                    Material material = new Material();
                    bool isNotCount = (string)mat.Attribute("Options") == "Project NotCount";
                    BuildBaseProperties(mat, material);

                    if ((string)mat.Attribute("Options") != "Project NotCount")
                    {
                        material.Quantity = isNotCount ? "" : (string)mat.Attribute("Quantity") ?? "";
                        material.BasePrice = isNotCount ? 0 : (decimal)ParseDouble((string)mat.Element("PriceBase")?.Attribute("Value"));
                        material.CurrPrice = isNotCount ? 0 : (decimal)ParseDouble((string)mat.Element("PriceCurr")?.Attribute("Value"));
                    }

                    // Не забыть что может быть Attribs="Deleted" "02-01-01и-4"
                    _position.Materials.Add(material);
                }
            }
            else if (_xPos?.Element("PriceCurr")?.Attributes("MT") != null)
            {
                Material material = new Material();
                BuildBaseProperties(_xPos, material);
                material.CurrPrice = (decimal)ParseDouble((string)_xPos.Element("PriceCurr")?.Attribute("MT"));
                _position.Materials.Add(material);
            }
        }

        public Position GetResult() 
        {
            Position result = _position;

            _position = new Position();
            return result; 
        }

        private static double ParseDouble(string value)
        {
            if (string.IsNullOrEmpty(value)) return 0;
            double.TryParse(value.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var result);
            return result;
        }
    }
}

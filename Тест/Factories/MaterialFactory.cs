using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Тест.Domain;

namespace Тест.Fabrics
{
    public class MaterialFactory : CalculationFactory
    {
        private readonly string _tagName = "Mat";
        public override List<IСalculation> GetCalculation(XmlDocument xmlDoc)
        {
            xmlDoc.GetElementsByTagName(_tagName);
            Material material = new Material();
            throw new NotImplementedException();
        }
    }
}

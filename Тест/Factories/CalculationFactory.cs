using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Тест.Domain;

namespace Тест.Fabrics
{
    public abstract class CalculationFactory
    {
        private readonly string? _tagName;
        public abstract List<IСalculation> GetCalculation(XmlDocument xmlDoc);
    }
}

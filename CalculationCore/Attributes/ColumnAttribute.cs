using System;

namespace CalculationCore.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnAttribute : Attribute
    {
        public string Header { get; }
        public int Order { get; }
        public string Format { get; set; }
        public ColumnAttribute(string header, int order = 0)
        {
            Header = header;
            Order = order;
            Format = Format;
        }
    }
}

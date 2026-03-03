using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Гранд_Смета.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnAttribute : Attribute
    {
        public string Header { get; }
        public int Order { get; }
        public ColumnAttribute(string header, int order = 0)
        {
            Header = header;
            Order = order;
        }
    }
}

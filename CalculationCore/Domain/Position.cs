using CalculationCore.Attributes;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace CalculationCore.Domain
{
    public class Position : ICalculation
    {
        [Column("Файл", 1)]
        public string FileName { get; set; }
        [Column("Локальный номер", 2, Format = "@")]
        public string LocNumber { get; set; }
        [Column("Наименование", 3)]
        public string Name { get; set; }
        [Column("Код", 4)]
        public string Code { get; set; }
        public string Chapter { get; set; }
        public string Quantity { get; set; }
        public string Unit { get; set; }
        [Column("Материал", 5)]
        public List<Material> Materials { get; set; } = new List<Material>();
    }
}

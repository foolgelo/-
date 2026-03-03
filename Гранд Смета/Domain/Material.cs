using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Гранд_Смета.Attributes;

namespace Гранд_Смета.Domain
{
    public class Material
    {
        [Column("Файл", 1)]
        public string fileName {  get; set; }
        [Column("Наименование", 2)]
        public string Name { get; set; }
        [Column("Код", 3)]
        public string Code { get; set; }
        [Column("Единицы измерения", 4)]
        public string Unit { get; set; }
        [Column("Количество", 5)]
        public decimal Quantity { get; set; }
        [Column("Цена", 6)]
        public decimal Price { get; set; }
    }
}

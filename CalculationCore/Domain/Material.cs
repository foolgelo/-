using CalculationCore.Attributes;

namespace CalculationCore.Domain
{
    public class Material
    {
        [Column("Файл", 1)]
        public string fileName {  get; set; }
        [Column("Номер сметы", 2, Format = "@")]
        public string LocNumber { get; set; }
        [Column("Наименование", 3)]
        public string Name { get; set; }
        [Column("Код", 4)]
        public string Code { get; set; }
        [Column("Единицы измерения", 5)]
        public string Unit { get; set; }
        [Column("Количество", 6)]
        public decimal Quantity { get; set; }
        [Column("Базовая цена", 7, Format = "#,##0.00 [$р.-119];-#,##0.00 [$р.-119]")]
        public decimal BasePrice { get; set; }
        [Column("Текущая цена", 8, Format = "#,##0.00 [$р.-119];-#,##0.00 [$р.-119]")]
        public decimal CurrPrice { get; set; }
    }
}

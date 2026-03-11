using CalculationCore.Attributes;

namespace CalculationCore.Domain
{
    public class Material : ICalculation
    {
        [Column("Наименование", 8)]
        public string Name { get; set; }
        [Column("Код", 9)]
        public string Code { get; set; }
        [Column("Единицы измерения", 10)]
        public string Unit { get; set; }
        [Column("Количество", 11)]
        public string Quantity { get; set; }
        [Column("Базовая цена", 12, Format = "#,##0.00 [$р.-119];-#,##0.00 [$р.-119]")]
        public decimal BasePrice { get; set; }
        [Column("Текущая цена", 13, Format = "#,##0.00 [$р.-119];-#,##0.00 [$р.-119]")]
        public decimal CurrPrice { get; set; }
    }
}

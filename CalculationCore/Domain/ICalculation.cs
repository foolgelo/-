using CalculationCore.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace CalculationCore.Domain
{
    public interface ICalculation
    {
        string Name { get; set; }
        string Code { get; set; }
        string Unit { get; set; }
    }
}

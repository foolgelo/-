using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Xunit;
using CalculationCore.Domain;
using CalculationCore.Factories;

namespace TestGrandSmeta
{
    public class MaterialFactoryTest
    {
        private const string MockXml = @"
            <Root>
                <NodeA>
                    <Mat Caption='Гвозди' Code='01' Units='т' Quantity='0,002'>
                        <PriceCurr Value='200000,00' />
                    </Mat>
                </NodeA>
                <NodeB>
                    <Mat Caption='Доски' Code='11' Units='м3' Quantity='0,263'>
                        <PriceBase Value='70000' />
                    </Mat>
                </NodeB>
            </Root>";

        [Fact]
        public void Descendants_ShouldFindAllNestedMaterials()
        {
            // Arrange
            XDocument doc = XDocument.Parse(MockXml);

            // Act
            var materials = doc.Descendants("Mat").ToList();

            // Assert
            Assert.Equal(2, materials.Count);
        }

        [Theory]
        [InlineData("0,002", 0.002)]
        [InlineData("1.5", 1.5)]
        [InlineData("", 0)]
        [InlineData(null, 0)]
        public void ParseDouble_ShouldHandleVariousFormats(string input, double expected)
        {
            // Если метод ParseDouble приватный, тестируем его через создание объекта
            XElement el = new XElement("Mat", new XAttribute("Quantity", input ?? ""));

            var material = MaterialFactory.CreateFromXElement(el, "MyFile", "12-51-j");
        }
    }
}

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Гранд_Смета.Domain;
using Гранд_Смета.Factories;
using Гранд_Смета.Repository;
using Гранд_Смета.Services;

namespace Гранд_Смета.UI
{
    public partial class Grand_Smeta
    {
        private void Grand_Smeta_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string path = @"C:\Олег\Документы\Газпром";

                gsfxRepository repo = new gsfxRepository(path);


                var materials = repo.LoadXMLFromGSFX()
                                    .SelectMany(doc => doc.Descendants("Mat")
                                    .Select(MaterialFactory.CreateFromXElement))
                                    .ToList();

                var service = new ExcelExportService();
                service.ExportToNewSheet(materials, "Материалы", "TableMaterials");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Ошибка: " + ex.Message);
            }
        }
    }
}

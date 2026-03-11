using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CalculationCore.Repository;
using CalculationCore.Factories;
using Гранд_Смета.Services;
using CalculationCore.Domain;

namespace Гранд_Смета.UI
{
    public partial class Grand_Smeta
    {
        private void Grand_Smeta_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Выберите папку";
                fbd.ShowNewFolderButton = false;

                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string path = fbd.SelectedPath;

                    try
                    {
                        gsfxRepository repo = new gsfxRepository(path);
                        IEnumerable<Position> materials = repo.GetAllMaterials();


                        ExcelExportService service = new ExcelExportService();
                        service.ExportMaterials(materials, "Материалы", "TableMaterials");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                }
            }

        }
    }
}

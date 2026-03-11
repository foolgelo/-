using CalculationCore.Domain;
using CalculationCore.Attributes;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Гранд_Смета;

namespace Гранд_Смета.Services
{
    public class ExcelExportService
    {
        /// <summary>
        /// Экспортирует только материалы, а также FileName и LocNumber позиции.
        /// </summary>
        public void ExportMaterials(IEnumerable<Position> positions, string sheetName, string tableName)
        {
            var app = Globals.ThisAddIn.Application;
            // Собираем все материалы с нужными полями позиции
            var rows = positions
                .Where(p => p.Materials != null)
                .SelectMany(p => p.Materials.Select(m => new
                {
                    FileName = p.FileName,
                    LocNumber = p.LocNumber,
                    Name = m.Name,
                    Code = m.Code,
                    Unit = m.Unit,
                    Quantity = m.Quantity,
                    BasePrice = m.BasePrice,
                    CurrPrice = m.CurrPrice
                }))
                .ToList();

            if (!rows.Any())
            {
                MessageBox.Show("Нет данных для выгрузки.");
                return;
            }

            // Определяем заголовки и форматы
            var headers = new[] { "Файл", "Локальный номер", "Наименование", "Код", "Ед.", "Кол-во", "Базовая цена", "Текущая цена" };
            var formats = new[] { "@", "@", "@", "@", "@", "0.###", "0.00", "0.00" };

            int rowCount = rows.Count;
            int colCount = headers.Length;

            object[,] data = new object[rowCount, colCount];
            for (int i = 0; i < rowCount; i++)
            {
                data[i, 0] = rows[i].FileName ?? "";
                data[i, 1] = rows[i].LocNumber ?? "";
                data[i, 2] = rows[i].Name ?? "";
                data[i, 3] = rows[i].Code ?? "";
                data[i, 4] = rows[i].Unit ?? "";
                data[i, 5] = rows[i].Quantity ?? "";
                data[i, 6] = rows[i].BasePrice;
                data[i, 7] = rows[i].CurrPrice;
            }

            Worksheet sheet = GetOrCreateSheet(app, sheetName);
            sheet.Activate();

            app.ScreenUpdating = false;
            try
            {
                sheet.Cells.Clear();

                Range startRange = sheet.Range["A1"].Resize[rowCount + 1, colCount];
                ListObject lo = sheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, startRange, Type.Missing, XlYesNoGuess.xlYes);
                lo.Name = tableName;

                // Заголовки
                lo.HeaderRowRange.Value2 = headers;

                // Форматы
                for (int j = 0; j < colCount; j++)
                {
                    if (!string.IsNullOrEmpty(formats[j]))
                    {
                        Range colRange = lo.DataBodyRange.Columns[j + 1];
                        colRange.NumberFormat = formats[j];
                    }
                }

                lo.DataBodyRange.Value2 = data;
                sheet.Columns.AutoFit();
                app.ActiveWorkbook.RefreshAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при выгрузке: {ex.Message}");
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private Worksheet GetOrCreateSheet(Microsoft.Office.Interop.Excel.Application app, string name)
        {
            Worksheet sheet = app.Sheets.Cast<Worksheet>()
                .FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

            if (sheet == null)
            {
                sheet = (Worksheet)app.Sheets.Add(After: app.Sheets[app.Sheets.Count]);
                sheet.Name = name;
            }
            return sheet;
        }
    }
}
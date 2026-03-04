using CalculationCore.Attributes; // Ваша библиотека с атрибутами
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Гранд_Смета;

namespace Гранд_Смета.Services
{
    public class ExcelExportService
    {
        public void ExportToNewSheet<T>(IEnumerable<T> items, string sheetName, string tableName)
        {
            var app = Globals.ThisAddIn.Application;
            var itemList = items.ToList();

            if (!itemList.Any())
            {
                MessageBox.Show("Нет данных для выгрузки.");
                return;
            }

            // 1. Подготовка листа (создаем или очищаем существующий)
            Worksheet sheet = GetOrCreateSheet(app, sheetName);
            sheet.Activate();

            // 2. Сбор метаданных через Reflection (используем наш атрибут с Order)
            var props = typeof(T).GetProperties()
                .Where(p => Attribute.IsDefined(p, typeof(ColumnAttribute)))
                .Select(p => new
                {
                    Property = p,
                    Attr = (ColumnAttribute)Attribute.GetCustomAttribute(p, typeof(ColumnAttribute))
                })
                .OrderBy(x => x.Attr.Order)
                .ToList();

            int rows = itemList.Count;
            int cols = props.Count;

            // 3. Подготовка двумерного массива данных (быстрее, чем писать в каждую ячейку)
            object[,] data = new object[rows, cols];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    data[i, j] = props[j].Property.GetValue(itemList[i]) ?? string.Empty;
                }
            }

            app.ScreenUpdating = false;
            try
            {
                // Очищаем лист перед новой выгрузкой
                sheet.Cells.Clear();

                // 4. Создаем "Умную таблицу" (ListObject)
                // Резервируем место: заголовки (A1) + данные
                Range startRange = sheet.Range["A1"].Resize[rows + 1, cols];
                ListObject lo = sheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, startRange, Type.Missing, XlYesNoGuess.xlYes);
                lo.Name = tableName;

                // Заполняем заголовки
                string[] headers = props.Select(p => p.Attr.Header).ToArray();
                lo.HeaderRowRange.Value2 = headers;

                // 5. ПРИМЕНЯЕМ ФОРМАТЫ (До вставки данных!)
                // Это критично для кодов (чтобы не стали датами)
                for (int j = 0; j < cols; j++)
                {
                    if (!string.IsNullOrEmpty(props[j].Attr.Format))
                    {
                        // Выделяем колонку данных (без заголовка)
                        Range colRange = lo.DataBodyRange.Columns[j + 1];
                        colRange.NumberFormat = props[j].Attr.Format;
                    }
                }

                // 6. Вставляем массив данных целиком
                lo.DataBodyRange.Value2 = data;

                // Финальные штрихи
                sheet.Columns.AutoFit();

                // 7. Обновляем Power Query (он увидит таблицу по имени)
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
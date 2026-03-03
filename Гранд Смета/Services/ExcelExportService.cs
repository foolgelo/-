using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Гранд_Смета.Attributes;

namespace Гранд_Смета.Services
{
    public class ExcelExportService
    {
        public void ExportToNewSheet<T>(IEnumerable<T> items, string sheetName, string tableName)
        {
            var app = Globals.ThisAddIn.Application;

            // 1. Получаем или создаем лист
            Worksheet sheet = GetOrCreateSheet(app, sheetName);
            sheet.Activate(); // Переключаемся на него

            // 2. Получаем метаданные из атрибутов (Reflection)
            var props = typeof(T).GetProperties()
                .Where(p => p.IsDefined(typeof(ColumnAttribute), false))
                .Select(p => new {
                    Property = p,
                    Attr = p.GetCustomAttribute<ColumnAttribute>()
                })
                .OrderBy(x => x.Attr.Order)
                .ToList();

            if (!props.Any()) return;

            // 3. Формируем данные
            var itemList = items.ToList();
            string[] headers = props.Select(x => x.Attr.Header).ToArray();
            object[,] data = new object[itemList.Count, props.Count];

            for (int i = 0; i < itemList.Count; i++)
            {
                for (int j = 0; j < props.Count; j++)
                {
                    data[i, j] = props[j].Property.GetValue(itemList[i]) ?? string.Empty;
                }
            }

            // 4. Выгружаем (Bulk Update)
            app.ScreenUpdating = false;
            try
            {
                // Очищаем всё старое на листе перед новой выгрузкой
                sheet.Cells.ClearContents();

                // Создаем таблицу (ListObject) начиная с ячейки A1
                Range startRange = sheet.Range["$A$1"].Resize[itemList.Count + 1, props.Count];
                ListObject lo = sheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, startRange, Type.Missing, XlYesNoGuess.xlYes);
                lo.Name = tableName;

                // Вставляем заголовки и данные
                lo.HeaderRowRange.Value2 = headers;
                if (itemList.Count > 0)
                {
                    lo.DataBodyRange.Value2 = data;
                }

                // Автоподбор ширины колонок
                sheet.Columns.AutoFit();

                // Обновляем Power Query
                app.ActiveWorkbook.RefreshAll();
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private Worksheet GetOrCreateSheet(Application app, string name)
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

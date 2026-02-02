using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ExcelAddIn1.Core.Models;
using ExcelAddIn1.Core.Services;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public partial class OnClick
    {
        private void OnClick_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void BtnExport_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var sheet = (Excel.Worksheet)app.ActiveWorkbook.ActiveSheet;

            int headerRow = 1;
            int maxColumns = 100;

            // 1. Вызначаем загалоўкі
            var headers = Enumerable.Range(1, maxColumns)
                .Select(c => sheet.Cells[headerRow, c].Value2)
                .TakeWhile(v => v != null && !string.IsNullOrWhiteSpace($"{v}"))
                .Select(v => $"{v}")
                .ToList();

            int columnCount = headers.Count;
            System.Diagnostics.Debug.WriteLine($"Columns: {columnCount}");

            if (columnCount == 0)
                return;

            // 2. Вызначаем дыяпазон дадзеных
            int startRow = headerRow + 1;
            int endRow = startRow + 1000; 

            var range = sheet.Range[
                sheet.Cells[startRow, 1],
                sheet.Cells[endRow, columnCount]
            ];

            // 3. АДЗІН COM-выклік
            var values = (object[,])range.Value2;

            // 4. Апрацоўка ў .NET (без COM)
            for (int r = 1; r <= values.GetLength(0); r++)
            {
                var rowValues = new List<string>();
                bool isEmptyRow = true;

                for (int c = 1; c <= values.GetLength(1); c++)
                {
                    var cell = values[r, c];
                    string text = $"{cell}";

                    if (!string.IsNullOrWhiteSpace(text))
                        isEmptyRow = false;

                    rowValues.Add(text);
                }

                if (isEmptyRow)
                    break;

                System.Diagnostics.Debug.WriteLine(string.Join("\t", rowValues));
            }

            var people = new List<Person>();

            for (int r = 1; r <= values.GetLength(0); r++)
            {
                var row = Enumerable.Range(1, columnCount)
                    .Select(c => values[r, c])
                    .ToArray();

                if (row.All(v => string.IsNullOrWhiteSpace($"{v}")))
                    break;

                var mapper = new PersonMapper();
                var person = mapper.MapRowToPerson(row);
                people.Add(person);
            }

            var exporter = new JsonExportService();

            string filePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "people.json"
            );

            exporter.ExportToFile(people, filePath);

            System.Windows.Forms.MessageBox.Show(
                $"JSON exported to:\n{filePath}",
                "Export complete"
            );

        }

    }

}

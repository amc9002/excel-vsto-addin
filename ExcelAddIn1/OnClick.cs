using Microsoft.Office.Tools.Ribbon;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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
            var headers = Enumerable
                .Range(1, 100) // верхняя мяжа (разумная)
                .Select(c => sheet.Cells[1, c].Value)
                .TakeWhile(v => v != null && !string.IsNullOrWhiteSpace($"{v}"))
                .Select(v => $"{v}")
                .ToList();

            int columnCount = headers.Count;


            System.Diagnostics.Debug.WriteLine($"Total columns: {columnCount}");

            // 2. Чытаем радкі пад загалоўкамі
            int row = headerRow + 1;

            while (true)
            {
                bool isEmptyRow = Enumerable
                    .Range(1, columnCount)
                    .Select(c => sheet.Cells[row, c].Value)
                    .All(v => v == null || string.IsNullOrWhiteSpace($"{v}"));


                if (isEmptyRow)
                {
                    break; // дадзеныя скончыліся
                }

                // Выводзім радок у Debug
                System.Diagnostics.Debug.Write($"Row {row}: ");

                for (row = 2; ; row++)
                {
                    var values = Enumerable.Range(1, headers.Count)
                        .Select(c => sheet.Cells[row, c].Value)
                        .ToList();

                    if (values.All(v => string.IsNullOrWhiteSpace($"{v}")))
                        break;

                    Debug.WriteLine(string.Join("\t", values));
                }


                System.Diagnostics.Debug.WriteLine("");
                row++;
            }
        }



    }

}

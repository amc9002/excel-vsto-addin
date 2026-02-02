using Microsoft.Office.Tools.Ribbon;

using System;
using System.Collections.Generic;
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
            int col = 1;

            // 1. Вызначаем калёнкі па загалоўках
            while (true)
            {
                var headerValue = sheet.Cells[headerRow, col].Value;

                if (headerValue == null || string.IsNullOrWhiteSpace($"{headerValue}"))
                {
                    break;
                }

                col++;
            }

            int columnCount = col - 1;

            System.Diagnostics.Debug.WriteLine($"Total columns: {columnCount}");

            // 2. Чытаем радкі пад загалоўкамі
            int row = headerRow + 1;

            while (true)
            {
                bool isEmptyRow = true;

                for (int c = 1; c <= columnCount; c++)
                {
                    var cellValue = sheet.Cells[row, c].Value;

                    if (cellValue != null && !string.IsNullOrWhiteSpace($"{cellValue}"))
                    {
                        isEmptyRow = false;
                        break;
                    }
                }

                if (isEmptyRow)
                {
                    break; // дадзеныя скончыліся
                }

                // Выводзім радок у Debug
                System.Diagnostics.Debug.Write($"Row {row}: ");

                for (int c = 1; c <= columnCount; c++)
                {
                    var cellValue = sheet.Cells[row, c].Value;
                    System.Diagnostics.Debug.Write($"{cellValue}\t");
                }

                System.Diagnostics.Debug.WriteLine("");
                row++;
            }
        }



    }

}

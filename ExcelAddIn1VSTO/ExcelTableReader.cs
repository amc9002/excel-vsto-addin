using System.Collections.Generic;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public class ExcelTableReader
    {
        public List<object[]> ReadTable(Excel.Worksheet sheet)
        {
            int headerRow = 1;
            int maxColumns = 100;

            var headers = Enumerable.Range(1, maxColumns)
                .Select(c => sheet.Cells[headerRow, c].Value2)
                .TakeWhile(v => v != null && !string.IsNullOrWhiteSpace($"{v}"))
                .ToList();

            if (headers.Count == 0)
                return new List<object[]>();

            int startRow = headerRow + 1;
            int endRow = startRow + 1000;

            var range = sheet.Range[
                sheet.Cells[startRow, 1],
                sheet.Cells[endRow, headers.Count]
            ];

            var values = (object[,])range.Value2;

            var rows = new List<object[]>();

            for (int r = 1; r <= values.GetLength(0); r++)
            {
                var row = Enumerable.Range(1, headers.Count)
                    .Select(c => values[r, c])
                    .ToArray();

                if (row.All(v => string.IsNullOrWhiteSpace($"{v}")))
                    break;

                rows.Add(row);
            }

            return rows;
        }
    }

}

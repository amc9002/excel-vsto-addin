using System.Collections.Generic;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    /// <summary>
    /// Provides functionality to extract tabular data from an Excel worksheet.
    /// </summary>
    public class ExcelTableReader
    {
        /// <summary>
        /// Reads a table from the worksheet by detecting headers and capturing rows until an empty line is encountered.
        /// </summary>
        /// <param name="sheet">The worksheet to read from.</param>
        /// <returns>A list of object arrays, where each array represents a row of data.</returns>
        public List<object[]> ReadTable(Excel.Worksheet sheet)
        {
            // Define scanning constraints
            int headerRow = 1;
            int maxColumns = 100;

            // 1. Identify headers: scan the first row until the first empty cell
            var headers = Enumerable.Range(1, maxColumns)
                .Select(c => sheet.Cells[headerRow, c].Value2)
                .TakeWhile(v => v != null && !string.IsNullOrWhiteSpace($"{v}"))
                .ToList();

            if (headers.Count == 0)
                return new List<object[]>();

            // 2. Define the data range (from row 2 down to 1000 for performance)
            int startRow = headerRow + 1;
            int endRow = startRow + 1000;

            // Fetching the entire range at once is much faster than reading cell-by-cell
            var range = sheet.Range[
                sheet.Cells[startRow, 1],
                sheet.Cells[endRow, headers.Count]
            ];

            // 3. Convert COM range to a 2D object array (Note: Excel arrays are 1-based)
            var values = (object[,])range.Value2;
            var rows = new List<object[]>();

            // 4. Iterate through the array and stop at the first completely empty row
            for (int r = 1; r <= values.GetLength(0); r++)
            {
                var row = Enumerable.Range(1, headers.Count)
                    .Select(c => values[r, c])
                    .ToArray();

                // Check if the row contains any meaningful data
                if (row.All(v => string.IsNullOrWhiteSpace($"{v}")))
                    break;

                rows.Add(row);
            }

            return rows;
        }
    }

}

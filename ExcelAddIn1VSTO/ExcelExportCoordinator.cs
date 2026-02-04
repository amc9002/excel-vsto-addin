using ExcelAddIn1.Core.Models;
using ExcelAddIn1.Core.Services;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    /// <summary>
    /// Orchestrates the export process by reading data from Excel and invoking appropriate export services.
    /// </summary>
    public class ExcelExportCoordinator
    {

        /// <summary>
        /// Reads the table from the worksheet and exports it to JSON format.
        /// </summary>
        /// <param name="sheet">The source Excel worksheet.</param>
        /// <returns>The path to the generated file, or null if export failed.</returns>
        public string ExportToJson(Excel.Worksheet sheet) => ExecuteExport(sheet, new JsonExportService(), "people.json");

        /// <summary>
        /// Reads the table from the worksheet and exports it to CSV format.
        /// </summary>
        /// <param name="sheet">The source Excel worksheet.</param>
        /// <returns>The path to the generated file, or null if export failed.</returns>
        public string ExportToCsv(Excel.Worksheet sheet) => ExecuteExport(sheet, new CsvExportService(), "people.csv");
        
        /// <summary>
        /// Reads the table from the worksheet and exports it to XBRL format.
        /// </summary>
        /// <param name="sheet">The source Excel worksheet.</param>
        /// <returns>The path to the generated file, or null if export failed.</returns>
        public string ExportToXbrl(Excel.Worksheet sheet) => ExecuteExport(sheet, new XbrlExportService(), "people.xbrl");


        // ---------------- private ----------------

        /// <summary>
        /// Reads data from the Excel worksheet and maps it to a list of Person objects.
        /// </summary>
        /// <param name="sheet">The Excel worksheet containing the source data.</param>
        /// <returns>A list of mapped Person entities.</returns>
        private List<Person> ReadTable(Excel.Worksheet sheet)
        {
            var reader = new ExcelTableReader();
            var rows = reader.ReadTable(sheet);

            var mapper = new PersonMapper();
            return rows
                .Select(row => mapper.MapRowToPerson(row))
                .ToList();
        }

        /// <summary>
        /// Constructs a full file path by combining the user's Desktop folder with the given file name.
        /// </summary>
        /// <param name="fileName">The name of the file (including extension).</param>
        /// <returns>A string representing the absolute path on the Desktop.</returns>
        private string GetDefaultPath(string fileName)
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                fileName
            );
        }

        /// <summary>
        /// Generic execution method that handles data reading, path resolution, and service invocation.
        /// </summary>
        private string ExecuteExport(Excel.Worksheet sheet, BaseExportService service, string fileName)
        {
            var data = ReadTable(sheet);
            if (data.Count == 0) return null;

            string path = GetDefaultPath(fileName);
            service.Export(data, path);
            return path;
        }
    }
}

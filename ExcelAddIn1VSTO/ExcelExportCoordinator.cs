using ExcelAddIn1.Core.Services;

using System;
using System.IO;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public class ExcelExportCoordinator
    {
        public string ExportPeopleToJson(Excel.Worksheet sheet)
        {
            var reader = new ExcelTableReader();
            var rows = reader.ReadTable(sheet);

            var mapper = new PersonMapper();
            var people = rows
                .Select(row => mapper.MapRowToPerson(row))
                .ToList();

            var exporter = new JsonExportService();

            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "people.json"
            );

            exporter.ExportToFile(people, path);
            return path;
        }
    }

}

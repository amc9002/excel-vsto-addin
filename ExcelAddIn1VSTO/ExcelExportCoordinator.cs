using ExcelAddIn1.Core.Models;
using ExcelAddIn1.Core.Services;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public class ExcelExportCoordinator
    {
        public string ExportToJson(Excel.Worksheet sheet)
        {
            var people = ReadTable(sheet);

            var exporter = new JsonExportService();
            string path = GetDefaultPath("people.json");

            exporter.ExportToFile(people, path);
            return path;
        }

        public string ExportToCsv(Excel.Worksheet sheet)
        {
            var people = ReadTable(sheet);

            var exporter = new CsvExportService();
            string path = GetDefaultPath("people.csv");

            exporter.Export(people, path);
            return path;
        }

        public void ExportToXbrl(Excel.Worksheet sheet)
        {
            // 1. Чытаем табліцу (выкарыстоўваем твой метад ReadTable)
            var data = ReadTable(sheet);

            if (data.Count == 0)
            {
                MessageBox.Show("Табліца пустая ці не знойдзены загалоўкі.");
                return;
            }

            // 2. Выклікаем Core-сэрвіс (цяпер мы яго створым)
            var xbrlService = new ExcelAddIn1.Core.Services.XbrlExportService();
            string xbrlContent = xbrlService.CreateXbrl(data);

            // 3. Захоўваем вынік (пакуль проста пакажам, што атрымалася)
            MessageBox.Show(xbrlContent, "Generated XBRL Preview");
        }

        // ---------------- private ----------------

        private List<Person> ReadTable(Excel.Worksheet sheet)
        {
            var reader = new ExcelTableReader();
            var rows = reader.ReadTable(sheet);

            var mapper = new PersonMapper();
            return rows
                .Select(row => mapper.MapRowToPerson(row))
                .ToList();
        }

        private string GetDefaultPath(string fileName)
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                fileName
            );
        }
    }
}

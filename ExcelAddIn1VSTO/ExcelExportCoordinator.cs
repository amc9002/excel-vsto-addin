using ExcelAddIn1.Core.Models;
using ExcelAddIn1.Core.Services;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public class ExcelExportCoordinator
    {
        public string ExportPeopleToJson(Excel.Worksheet sheet)
        {
            var people = ReadPeople(sheet);

            var exporter = new JsonExportService();
            string path = GetDefaultPath("people.json");

            exporter.ExportToFile(people, path);
            return path;
        }

        public string ExportPeopleToCsv(Excel.Worksheet sheet)
        {
            var people = ReadPeople(sheet);

            var exporter = new CsvExportService();
            string path = GetDefaultPath("people.csv");

            exporter.Export(people, path);
            return path;
        }

        public void ExportToXbrl()
        {
            // Пакуль што проста дыягностыка, каб праверыць ланцужок
            System.Windows.Forms.MessageBox.Show("Coordinator атрымаў каманду на XBRL!");
        }

        // ---------------- private ----------------

        private List<Person> ReadPeople(Excel.Worksheet sheet)
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

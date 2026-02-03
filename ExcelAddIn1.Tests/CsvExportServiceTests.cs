using ExcelAddIn1.Core.Models;
using ExcelAddIn1.Core.Services;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn1.Tests
{
    public class CsvExportServiceTests
    {
        [Fact]
        public void Export_Should_Create_Valid_Csv_File()
        {
            var data = new[]
            {
                new Person { Name = "Alice", Age = 30, City = "Minsk" },
                new Person { Name = "Bob, Jr.", Age = null, City = "NY" }
            };

            var path = Path.GetTempFileName();
            var service = new CsvExportService();

            service.Export(data, path);

            var text = File.ReadAllText(path);

            Assert.Contains("Alice,30,Minsk", text);
            Assert.Contains("\"Bob, Jr.\",,NY", text);
        }

    }
}

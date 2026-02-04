using ExcelAddIn1.Core.Models;
using ExcelAddIn1.Core.Services;

using System;

namespace ExcelAddIn1.Tests
{
    public class XbrlExportServiceTests
    {
        [Fact]
        public void Export_Should_Handle_ValidationErrors_Properly()
        {
            var data = new List<Person>();

            var person = new Person { Name = "Ivan" };

            person.ValidationErrors.Add("Error 1");
            person.ValidationErrors.Add("Error 2");

            data.Add(person);

            var service = new XbrlExportService();
            var path = Path.GetTempFileName();

            try
            {
                // 2. Act
                service.Export(data, path);
                var content = File.ReadAllText(path);

                // 3. Assert
                Assert.Contains("errors=\"Error 1, Error 2\"", content);
                Assert.Contains("/>", content);
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}

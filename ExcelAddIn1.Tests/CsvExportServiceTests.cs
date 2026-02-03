using ExcelAddIn1.Core.Models;
using ExcelAddIn1.Core.Services;
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

            var path = Path.Combine(
                Path.GetTempPath(),
                Guid.NewGuid() + ".csv"
            );
            var service = new CsvExportService();

            service.Export(data, path);

            var lines = File.ReadAllLines(path);

            Assert.Equal("Name,Age,City,ValidationErrors", lines[0]);
            Assert.Equal("Alice,30,Minsk,", lines[1]);
            Assert.Equal("\"Bob, Jr.\",,NY,", lines[2]);
        }
    }
}

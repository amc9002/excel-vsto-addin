using ExcelAddIn1.Core.Models;
using ExcelAddIn1.Core.Services;

namespace ExcelAddIn1.Tests
{
    public class JsonExportServiceTests
    {
        [Fact]
        public void Serialize_Should_Return_Valid_Json()
        {
            // ARRANGE
            var people = new List<Person>
        {
            new() { Name = "Alice", Age = 30, City = "Minsk" }
        };

            var service = new JsonExportService();

            // ACT
            var json = service.Serialize(people);

            // ASSERT
            Assert.Contains("\"Name\": \"Alice\"", json);
            Assert.Contains("\"Age\": 30", json);
            Assert.Contains("\"City\": \"Minsk\"", json);
        }

        [Fact]
        public void Serialize_Should_Return_Empty_Array_For_Empty_List()
        {
            var service = new JsonExportService();

            var json = service.Serialize(new List<Person>());

            Assert.Equal("[]", json.Trim());
        }

        [Fact]
        public void Serialize_Should_Throw_When_Data_Is_Null()
        {
            var service = new JsonExportService();

            Assert.Throws<ArgumentNullException>(() =>
                service.Serialize<Person>(null)
            );
        }

        [Fact]
        public void ExportToFile_Should_Create_File_With_Content()
        {
            var people = new List<Person>
            {
                new() { Name = "Alice", Age = 30, City = "Minsk" }
            };

            var service = new JsonExportService();
            var path = Path.GetTempFileName();

            try
            {
                service.ExportToFile(people, path);

                var content = File.ReadAllText(path);
                Assert.Contains("Alice", content);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Export_Should_Write_Empty_Array_When_Data_Is_Empty()
        {
            var service = new JsonExportService();
            var path = Path.GetTempFileName();

            service.ExportToFile(new List<Person>(), path);

            var json = File.ReadAllText(path);
            Assert.Equal("[]", json.Trim());
        }

        [Fact]
        public void Export_Should_Overwrite_Existing_File()
        {
            var path = Path.GetTempFileName();
            File.WriteAllText(path, "OLD");

            new JsonExportService().ExportToFile(
                new[] { new Person { Name = "Alice" } },
                path
            );

            var json = File.ReadAllText(path);
            Assert.Contains("Alice", json);
        }

    }
}

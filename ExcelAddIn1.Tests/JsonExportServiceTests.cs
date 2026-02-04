using ExcelAddIn1.Core.Models;
using ExcelAddIn1.Core.Services;

namespace ExcelAddIn1.Tests
{
    public class JsonExportServiceTests
    {
        [Fact]
        public void Export_Should_Save_Valid_Json_Structure()
        {
            // ARRANGE
            var service = new JsonExportService();
            var path = Path.GetTempFileName();
            var data = new List<Person>
            {
                new() { Name = "Alice", Age = 30, City = "Minsk" }
            };

            try
            {
                // ACT
                service.Export(data, path);

                // ASSERT
                var json = File.ReadAllText(path);
                Assert.Contains("\"Name\": \"Alice\"", json);
                Assert.Contains("\"Age\": 30", json);
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Export_Should_Write_Empty_Array_To_File_When_Data_Is_Empty()
        {
            // ARRANGE
            var service = new JsonExportService();
            var path = Path.GetTempFileName();
            var emptyList = new List<Person>();

            try
            {
                // ACT
                service.Export(emptyList, path);

                // ASSERT
                var json = File.ReadAllText(path).Trim();
                Assert.Equal("[]", json);
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Export_Should_Throw_ArgumentNullException_When_Data_Is_Null()
        {
            // ARRANGE
            var service = new JsonExportService();
            var path = "any_path.json";

            // ACT & ASSERT
            Assert.Throws<ArgumentNullException>(() =>
                service.Export<Person>(null, path));
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
                service.Export(people, path);

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

            service.Export(new List<Person>(), path);

            var json = File.ReadAllText(path);
            Assert.Equal("[]", json.Trim());
        }

        [Fact]
        public void Export_Should_Overwrite_Existing_File()
        {
            var path = Path.GetTempFileName();
            File.WriteAllText(path, "OLD");

            new JsonExportService().Export(
                new[] { new Person { Name = "Alice" } },
                path
            );

            var json = File.ReadAllText(path);
            Assert.Contains("Alice", json);
        }

    }
}

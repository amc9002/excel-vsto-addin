using ExcelAddIn1.Core.Services;
using Xunit;

namespace ExcelAddIn1.Tests
{
    public class PersonMapperTests
    {
        [Fact]
        public void MapRowToPerson_Should_Map_All_Fields()
        {
            // ARRANGE
            object[] row =
            {
            "Alice",
            "30",
            "Minsk"
            };

            var mapper = new PersonMapper();

            // ACT
            var person = mapper.MapRowToPerson(row);

            // ASSERT
            Assert.Equal("Alice", person.Name);
            Assert.Equal(30, person.Age);
            Assert.Equal("Minsk", person.City);
        }
    }

}

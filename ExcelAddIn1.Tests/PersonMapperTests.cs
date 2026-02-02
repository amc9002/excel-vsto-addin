using ExcelAddIn1.Core.Services;

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

        [Fact]
        public void MapRow_Should_Handle_Null_And_Empty_Values()
        {
            object[] row = { null, "", "Minsk" };

            var mapper = new PersonMapper();
            var person = mapper.MapRowToPerson(row);

            Assert.Null(person.Name);
            Assert.Null(person.Age);
            Assert.Equal("Minsk", person.City);
        }

        [Fact]
        public void MapRow_Should_Not_Throw_When_Row_Is_Short()
        {
            object[] row = { "Alice" };

            var mapper = new PersonMapper();
            var person = mapper.MapRowToPerson(row);

            Assert.Equal("Alice", person.Name);
            Assert.Null(person.Age);
            Assert.Null(person.City);
        }
    }
}

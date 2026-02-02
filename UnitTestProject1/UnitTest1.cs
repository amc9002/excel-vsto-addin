using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelAddIn1.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;


[TestClass]
public class PersonMappingTests
{
    [TestMethod]
    public void MapRowToPerson_Should_Map_All_Fields()
    {
        // ---------- ARRANGE ----------
        object[] row =
        {
            "Alice",
            "30",
            "Minsk"
        };

        var mapper = new PersonMapper();
        // або твой клас, дзе MapRowToPerson

        // ---------- ACT ----------
        var person = mapper.MapRowToPerson(row);

        // ---------- ASSERT ----------
        Assert.AreEqual("Alice", person.Name);
        Assert.AreEqual(30, person.Age);
        Assert.AreEqual("Minsk", person.City);
    }
}


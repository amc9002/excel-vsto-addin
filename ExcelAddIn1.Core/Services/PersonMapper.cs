using ExcelAddIn1.Core.Models;

namespace ExcelAddIn1.Core.Services
{
    public class PersonMapper
    {
        public Person MapRowToPerson(object[] row)
        {
            return new Person
            {
                Name = row[0]?.ToString(),
                Age = int.TryParse(row[1]?.ToString(), out var age) ? age : (int?)null,
                City = row[2]?.ToString()
            };
        }
    }
}


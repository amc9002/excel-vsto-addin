using ExcelAddIn1.Core.Models;

using System;

namespace ExcelAddIn1.Core.Services
{
    public class PersonMapper
    {
        public Person MapRowToPerson(object[] row)
        {
            if (row == null)
                throw new ArgumentNullException(nameof(row)); 

            Person person = new Person();

            person.Name = row.Length > 0 ? row[0]?.ToString() : null;

            if (row.Length > 1 && int.TryParse(row[1]?.ToString(), out var age))
                person.Age = age;

            person.City = row.Length > 2 ? row[2]?.ToString() : null;

            return person;

        }
    }
}


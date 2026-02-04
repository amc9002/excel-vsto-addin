using ExcelAddIn1.Core.Models;

using Newtonsoft.Json.Linq;

using System;

namespace ExcelAddIn1.Core.Services
{
    public class PersonMapper
    {
        public Person MapRowToPerson(object[] row)
        {
            var person = new Person();

            person.Name = GetString(row, 0);

            if (TryGetInt(row, 1, out int age))
                person.Age = age;

            if (age < 0 || age > 130)
            {
                person.ValidationErrors.Add("Invalid age");
            }

            person.City = GetString(row, 2);

            return person;
        }

        private string GetString(object[] row, int index)
        {
            if (row == null || index >= row.Length)
                return null;

            var value = row[index];
            var text = $"{value}".Trim();

            return string.IsNullOrWhiteSpace(text) ? null : text;
        }

        private bool TryGetInt(object[] row, int index, out int value)
        {
            value = default;

            if (row == null || index >= row.Length)
                return false;

            var raw = row[index];
            if (raw == null)
                return false;

            var text = raw.ToString().Trim();
            if (string.IsNullOrEmpty(text))
                return false;

            if (!int.TryParse(text, out value))
                return false;

            return true;
        }
    }

}



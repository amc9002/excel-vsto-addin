using ExcelAddIn1.Core.Models;

using Newtonsoft.Json.Linq;

using System;

namespace ExcelAddIn1.Core.Services
{
    /// <summary>
    /// Responsible for transforming raw data arrays (from Excel) into strongly-typed Person objects.
    /// Includes basic validation and data type conversion.
    /// </summary>
    public class PersonMapper
    {
        /// <summary>
        /// Maps an array of objects to a Person instance, performing type conversion and age validation.
        /// </summary>
        /// <param name="row">An array of objects representing a row from Excel (expected: [Name, Age, City]).</param>
        /// <returns>A Person object with mapped properties and potential validation errors.</returns>
        public Person MapRowToPerson(object[] row)
        {
            var person = new Person();

            // Map Name from the first column
            person.Name = GetString(row, 0);

            // Attempt to parse Age from the second column
            if (TryGetInt(row, 1, out int age))
                person.Age = age;

            // Business Rule: Validate human age range. 
            // Note: Even if invalid, the value is preserved for user correction.
            if (age < 0 || age > 130)
            {
                person.ValidationErrors.Add("Invalid age");
            }

            // Map City from the third column
            person.City = GetString(row, 2);

            return person;
        }

        /// <summary>
        /// Safely extracts and trims a string value from the object array.
        /// </summary>
        private string GetString(object[] row, int index)
        {
            if (row == null || index >= row.Length)
                return null;

            var value = row[index];
            var text = $"{value}".Trim();

            return string.IsNullOrWhiteSpace(text) ? null : text;
        }

        /// <summary>
        /// Attempts to parse an object value into an integer. Returns false if parsing fails or value is null.
        /// </summary>
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



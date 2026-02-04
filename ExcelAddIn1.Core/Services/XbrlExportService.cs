using ExcelAddIn1.Core.Models;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelAddIn1.Core.Services
{
    /// <summary>
    /// Service for exporting data into XBRL (Extensible Business Reporting Language) format.
    /// Uses reflection to map object properties to XML attributes.
    /// </summary>
    public class XbrlExportService : BaseExportService
    {
        /// <summary>
        /// Generates the XBRL XML string representation of the provided data.
        /// </summary>
        /// <param name="data">List of Person objects to be converted to XML.</param>
        /// <returns>A formatted XBRL XML string.</returns>
        public string CreateXbrl(List<Person> data)
        {
            StringBuilder sb = new StringBuilder();

            // 1. Add XML declaration and root element with namespace
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.AppendLine("<xbrl xmlns=\"http://www.xbrl.org/2003/instance\">");

            foreach (var p in data)
            {
                // Start entry element with indentation for better readability
                sb.Append("  <entry ");

                // 2. Use Reflection to iterate through all properties of the Person class
                PropertyInfo[] properties = p.GetType().GetProperties();
                foreach (var prop in properties)
                {
                    string name = prop.Name.ToLower();
                    object value = prop.GetValue(p);

                    // 3. Special handling for Validation Errors (collection of strings)
                    if (name == "validationerrors")
                    {
                        var errors = value as List<string>;
                        if (errors != null && errors.Count > 0)
                        {
                            // Join all errors into a single attribute string
                            sb.Append($"errors=\"{string.Join(", ", errors)}\" ");
                        }
                        continue;
                    }

                    // 4. Filter properties: 
                    // We only want simple types or Nullable value types.
                    // We skip complex generic collections/lists (handled above or ignored).
                    if (prop.PropertyType.IsGenericType &&
                        prop.PropertyType.GetGenericTypeDefinition() != typeof(Nullable<>) &&
                        name != "validationerrors")
                        continue;

                    // 5. Append property as an XML attribute: name="value"
                    sb.Append($"{name}=\"{value}\" ");
                }

                // Close the self-closing entry tag and move to the next line
                sb.AppendLine("/>");
            }

            // 6. Close the root element
            sb.AppendLine("</xbrl>");
            return sb.ToString();
        }

        /// <summary>
        /// Implementation of the base Export method. Maps generic data to Person list and saves to file.
        /// </summary>
        public override void Export<T>(IEnumerable<T> data, string filePath)
        {
            // For XBRL, we specifically expect Person objects
            var people = data.Cast<Person>().ToList();
            string content = CreateXbrl(people);

            // SaveToFile is inherited from BaseExportService
            SaveToFile(content, filePath);
        }
    }
}

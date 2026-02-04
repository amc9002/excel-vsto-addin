using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAddIn1.Core.Services
{
    /// <summary>
    /// Service for exporting generic data collections into CSV format.
    /// Supports automatic header generation and complex type handling (like lists).
    /// </summary>
    public class CsvExportService : BaseExportService
    {
        /// <summary>
        /// Exports the data collection to a CSV file.
        /// </summary>
        /// <typeparam name="T">The type of objects in the collection.</typeparam>
        /// <param name="data">Collection of objects to export.</param>
        /// <param name="filePath">Target file path.</param>
        public override void Export<T>(IEnumerable<T> data, string filePath)
        {
            // 1. Get all readable properties of the class using Reflection
            var properties = typeof(T)
                .GetProperties()
                .Where(p => p.CanRead)
                .ToArray();

            // 2. Generate Header Row: extract property names and escape them
            var separator = ",";
            var lines = new List<string>
            {
                // Header
                string.Join(
                separator,
                properties.Select(p => Escape(p.Name, separator))
                )
            };

            // 3. Process Data Rows
            foreach (var item in data)
            {
                var values = properties.Select(p =>
                {
                    var value = p.GetValue(item);

                    string text = string.Empty;

                    // Handle different value types for proper string representation
                    switch (value)
                    {
                        case null: break;

                        // Special case: if the property is a list (like ValidationErrors), 
                        // join its elements into a single semicolon-separated string
                        case IEnumerable<string> list:
                            {
                                text = list.Any()
                                    ? string.Join("; ", list)
                                    : string.Empty;
                                break;
                            }
                        default:
                            text = value.ToString(); break;
                    }

                    // Escape the resulting text to ensure CSV integrity
                    return Escape(text, separator);

                });

                lines.Add(string.Join(separator, values));
            }

            // 4. Combine all lines and save using the base service method
            var finalContent = string.Join(Environment.NewLine, lines);
            SaveToFile(finalContent, filePath);
        }

        /// <summary>
        /// Escapes special characters in a CSV field according to RFC 4180.
        /// Wraps the value in quotes if it contains separators, quotes, or newlines.
        /// </summary>
        private string Escape(string value, string separator)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;

            // Check if the value contains characters that require quoting
            if (value.Contains(separator) || value.Contains("\"") || value.Contains("\n"))
            {
                // Double up any existing quotes and wrap the whole string in quotes
                return $"\"{value.Replace("\"", "\"\"")}\"";
            }

            return value;
        }
    }
}


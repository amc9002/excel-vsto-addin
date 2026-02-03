using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelAddIn1.Core.Services
{
    public class CsvExportService
    {
        public void Export<T>(
            IEnumerable<T> data,
            string filePath,
            string separator = ","
        )
        {
            var properties = typeof(T)
                .GetProperties()
                .Where(p => p.CanRead)
                .ToArray();

            var lines = new List<string>
            {
                // Header
                string.Join(
                separator,
                properties.Select(p => Escape(p.Name, separator))
            )
            };

            // Rows
            foreach (var item in data)
            {
                var values = properties.Select(p =>
                {
                    var value = p.GetValue(item);

                string text = string.Empty;
                    switch(value)
                    {
                        case null: break;
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
                    
                    return Escape(text, separator);

                });

                lines.Add(string.Join(separator, values));
            }

            File.WriteAllLines(filePath, lines, Encoding.UTF8);
        }

        private string Escape(string value, string separator)
        {
            if (value.Contains(separator) || value.Contains("\"") || value.Contains("\n"))
            {
                value = value.Replace("\"", "\"\"");
                return $"\"{value}\"";
            }

            return value;
        }
    }
}


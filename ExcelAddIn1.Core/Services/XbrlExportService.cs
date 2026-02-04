using ExcelAddIn1.Core.Models;

using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ExcelAddIn1.Core.Services
{
    public class XbrlExportService
    {
        public string CreateXbrl(List<Person> data)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sb.AppendLine("<xbrl xmlns=\"http://www.xbrl.org/2003/instance\">");

            foreach (var p in data)
            {
                sb.Append("  <entry "); // Дадаем прабелы для адступу

                PropertyInfo[] properties = p.GetType().GetProperties();
                foreach (var prop in properties)
                {
                    string name = prop.Name.ToLower();
                    object value = prop.GetValue(p);

                    if (name == "validationerrors")
                    {
                        var errors = value as List<string>;
                        if (errors != null && errors.Count > 0)
                        {
                            sb.Append($"errors=\"{string.Join(", ", errors)}\" ");
                        }
                        continue;
                    }

                    // Пакідаем Nullable і простыя тыпы
                    if (prop.PropertyType.IsGenericType &&
                        prop.PropertyType.GetGenericTypeDefinition() != typeof(Nullable<>) &&
                        name != "validationerrors")
                        continue;

                    sb.Append($"{name}=\"{value}\" ");
                }

                sb.AppendLine("/>"); // Закрываем entry і пераходзім на новы радок
            }

            sb.AppendLine("</xbrl>"); // Закрываем агульны тэг пасля цыкла
            return sb.ToString();
        }
    }
}

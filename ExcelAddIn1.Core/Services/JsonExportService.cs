using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelAddIn1.Core.Services
{
    public class JsonExportService
    {
        public string Serialize<T>(IEnumerable<T> data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            return JsonConvert.SerializeObject(
                data,
                Formatting.Indented
            );
        }

        public void ExportToFile<T>(IEnumerable<T> data, string filePath)
        {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));

            var json = Serialize(data);
            File.WriteAllText(filePath, json);
        }
    }


}

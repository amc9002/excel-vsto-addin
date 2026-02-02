using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

public class JsonExportService
{
    public void Export<T>(IEnumerable<T> data, string filePath)
    {
        string json = JsonConvert.SerializeObject(
            data,
            Formatting.Indented
        );

        File.WriteAllText(filePath, json);
    }
}

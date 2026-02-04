using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelAddIn1.Core.Services
{
    /// <summary>
    /// Service for exporting data into JSON format using the Newtonsoft.Json library.
    /// Provides a high-level abstraction over standard serialization.
    /// </summary>
    public class JsonExportService : BaseExportService
    {
        /// <summary>
        /// Serializes the data collection into a formatted JSON string and saves it to a file.
        /// </summary>
        /// <typeparam name="T">The type of data objects in the collection.</typeparam>
        /// <param name="data">The collection of objects to be serialized.</param>
        /// <param name="filePath">The destination file path.</param>
        /// <exception cref="ArgumentNullException">Thrown if data or filePath is null.</exception>
        public override void Export<T>(IEnumerable<T> data, string filePath)
        {
            // Validate inputs before processing
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));

            // 1. Serialize the collection to a string with indented (pretty) formatting
            var json = JsonConvert.SerializeObject(data, Formatting.Indented);

            // 2. Save the resulting JSON string using the shared base method
            SaveToFile(json, filePath);
        }
    }
}

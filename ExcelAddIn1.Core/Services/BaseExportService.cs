
using System.Collections.Generic;

public interface ITextExportable
{
    void SaveToFile(string content, string filePath);
}

/// <summary>
/// Provides a base implementation for data export services.
/// </summary>
public abstract class BaseExportService
{
    /// <summary>
    /// Writes the provided string content to a file using UTF-8 encoding.
    /// </summary>
    /// <param name="content">The text content to be saved.</param>
    /// <param name="filePath">The full path where the file will be created.</param>
    public void SaveToFile(string content, string filePath)
    {
        System.IO.File.WriteAllText(filePath, content, System.Text.Encoding.UTF8);
    }

    /// <summary>
    /// When overridden in a derived class, exports the collection of data to a specified file.
    /// </summary>
    /// <typeparam name="T">The type of data objects in the collection.</typeparam>
    /// <param name="data">The collection of objects to export.</param>
    /// <param name="filePath">The destination file path.</param>
    public abstract void Export<T>(IEnumerable<T> data, string filePath);
}

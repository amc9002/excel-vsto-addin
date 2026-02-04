
namespace ExcelAddIn1.Core.Services
{
    /// <summary>
    /// Defines a contract for services that handle the final stage of data export 
    /// by persisting string content to the file system.
    /// </summary>
    public interface IExportService
    {
        /// <summary>
        /// Saves the generated content string to a specified file path.
        /// </summary>
        /// <param name="content">The formatted string (JSON, CSV, XBRL, etc.) to be written.</param>
        /// <param name="filePath">The absolute path where the file should be saved.</param>
        void Export(string content, string filePath);
    }
}

using Microsoft.Office.Core;

using System.Windows.Forms;

namespace ExcelAddIn1
{
    /// <summary>
    /// Handles the Ribbon interface for the Excel Add-in.
    /// Manages the custom UI and dispatches user actions to the export services.
    /// </summary>
    [System.Runtime.InteropServices.ComVisible(true)]
    public partial class AddInRibbon : IRibbonExtensibility
    {
        /// <summary>
        /// Loads the XML custom UI definition from the assembly resources.
        /// </summary>
        /// <param name="ribbonID">The ID of the ribbon to load (managed by Office).</param>
        /// <returns>The XML string representing the ribbon layout.</returns>
        public string GetCustomUI(string ribbonID)
        {
            try
            {
                // Loading the ribbon layout from Properties.Resources
                return System.Text.Encoding.UTF8.GetString(Properties.Resources.AddInRibbon);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Resource Load Error: " + ex.Message);
                return "";
            }
        }

        /// <summary>
        /// Callback method triggered when a user clicks any export button on the Ribbon.
        /// Uses the control ID to determine the target export format.
        /// </summary>
        /// <param name="control">The Ribbon control that triggered the event (btnExportJson, btnExportCsv, or btnExportXbrl).</param>
        public void OnTestClicked(IRibbonControl control)
        {
            // 1. Get reference to the Excel Application and Active Sheet
            var app = Globals.ThisAddIn.Application;
            var sheet = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveWorkbook.ActiveSheet;

            // 2. Initialize the coordinator to handle the data flow
            var coordinator = new ExcelExportCoordinator();
            string filePath = string.Empty;

            // 3. Dispatch action based on which button was clicked
            switch (control.Id)
            {
                case "btnExportJson":
                    filePath = coordinator.ExportToJson(sheet);
                    ShowSuccessMessage(filePath, "JSON");
                    break;

                case "btnExportCsv":
                    filePath = coordinator.ExportToCsv(sheet);
                    ShowSuccessMessage(filePath, "CSV");
                    break;

                case "btnExportXbrl":
                    filePath = coordinator.ExportToXbrl(sheet);
                    ShowSuccessMessage(filePath, "XBRL");
                    break;
            }
        }

        /// <summary>
        /// Helper method to display a consistent success notification to the user.
        /// </summary>
        private void ShowSuccessMessage(string path, string format)
        {
            if (!string.IsNullOrEmpty(path))
            {
                MessageBox.Show($"{format} exported to:\n{path}", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
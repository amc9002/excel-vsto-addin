using Microsoft.Office.Core;

using System.Windows.Forms;

namespace ExcelAddIn1
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public partial class AddInRibbon : IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonID)
        {
            // Гэта апошні шанец убачыць памылку
            try
            {
                return System.Text.Encoding.UTF8.GetString(Properties.Resources.AddInRibbon);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Памылка рэсурсу: " + ex.Message);
                return "";
            }
        }

        // Метад для кнопкі
        public void OnTestClicked(IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            var sheet = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveWorkbook.ActiveSheet;
            var coordinator = new ExcelExportCoordinator();
            string filePath = string.Empty;

            switch (control.Id)
            {
                case "btnExportJson":
                    filePath = coordinator.ExportToJson(sheet);
                    MessageBox.Show($"JSON exported to:\n{filePath}", "Export complete");
                    break;

                case "btnExportCsv":
                    filePath = coordinator.ExportToCsv(sheet);
                    MessageBox.Show($"CSV exported to:\n{filePath}", "Export complete");
                    break;

                case "btnExportXbrl":
                    coordinator.ExportToXbrl(sheet); 
                    break;
            }
        }
    }
}

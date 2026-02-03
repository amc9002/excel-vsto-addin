using ExcelAddIn1.Core.Models;
using ExcelAddIn1.Core.Services;

using Microsoft.Office.Tools.Ribbon;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public partial class OnClick
    {
        private void OnClick_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void BtnExportJSON_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var sheet = (Excel.Worksheet)app.ActiveWorkbook.ActiveSheet;

            var coordinator = new ExcelExportCoordinator();
            string filePath = coordinator.ExportPeopleToJson(sheet);

            MessageBox.Show(
                $"JSON exported to:\n{filePath}",
                "Export complete"
            );
        }

        private void BtnExportCSV_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var sheet = (Excel.Worksheet)app.ActiveWorkbook.ActiveSheet;

            var coordinator = new ExcelExportCoordinator();
            string filePath = coordinator.ExportPeopleToCsv(sheet);

            MessageBox.Show(
                $"CSV exported to:\n{filePath}",
                "Export complete"
            );
        }

    }
}

using Microsoft.Office.Tools.Ribbon;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public partial class OnClick
    {
        private void OnClick_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void BtnExport_Click(object sender, RibbonControlEventArgs e)
        {
            var sheet = (Excel.Worksheet)
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            sheet.Cells[1, 1].Value = "Hello from VSTO";
        }

    }

}

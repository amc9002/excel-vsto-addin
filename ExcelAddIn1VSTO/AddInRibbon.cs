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
        public void OnTestClicked(Microsoft.Office.Core.IRibbonControl control)
        {
            // Проста паказваем, што менавіта націснута
            MessageBox.Show($"Выбраны фармат: {control.Id}");

            var coordinator = new ExcelExportCoordinator();

            // Лагічны разгалiнавальнік
            switch (control.Id)
            {
                case "btnExportJson":
                    // Тут будзе выклік JSON логікі
                    break;
                case "btnExportCsv":
                    // Тут будзе выклік CSV логікі
                    break;
                case "btnExportXbrl":
                    coordinator.ExportToXbrl();
                    break;
            }
        }
    }
}

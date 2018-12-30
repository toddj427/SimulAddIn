using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace SimulationAddIn
{
    public partial class SimRibbon
    {
        ConfigDataBL myConfigDataBL = null;
        ConfigSheet myConfigSheet = new ConfigSheet();


        private void SimRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            myConfigDataBL = new ConfigDataBL();

        }

        private void btnModelPath_Click(object sender, RibbonControlEventArgs e)
        {
            myConfigDataBL.SetModelPath();
            myConfigSheet.UpdateModelPath(myConfigDataBL);
        }

        private void btnInitialize_Click(object sender, RibbonControlEventArgs e)
        {
            string title = "Warning";
            string message = "This command will delete all previously set data and delete/recreate configuration sheets.  Do you want to continue?";

            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.No)
            {
                return;
            }

            myConfigSheet.InitializeNewSheet();

        }
    }
}

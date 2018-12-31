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

        private void BtnInitialize_Click(object sender, RibbonControlEventArgs e)
        {
            string title = "Warning";
            string message = "This command will delete all previously set data and delete/recreate configuration sheets.  Do you want to continue?";

            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.No)
            {
                return;
            }

            myConfigSheet.CreateNewConfigSheet();

            // get permission, then set the model path.
            title = "Step 2";
            message = "Would you like to set the model path now?";
            buttons = MessageBoxButtons.YesNo;
            result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.Yes)
            {
                myConfigDataBL.SetModelPath();

            }

            // Open question: Is it ok to set the model name if we haven't decided on a path?
            // Maybe set things up to do both if the model already exists?

            // Set up the model name.
            title = "Step 3";
            message = "Would you like to set the model name?";
            buttons = MessageBoxButtons.YesNo;
            result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.Yes)
            {
                myConfigDataBL.SetModelFile();
            }

            // Display a "done" message
            title = "";
            message = "Initialization is complete.";
            buttons = MessageBoxButtons.OK;
            MessageBox.Show(message, title, buttons);
        }

        // Ask the user for the path to the model
        private void btnModelPath_Click(object sender, RibbonControlEventArgs e)
        {
            myConfigDataBL.SetModelPath();
        }

        // Ask the user for the model name or the model executable file
        private void BtnModelName_Click(object sender, RibbonControlEventArgs e)
        {
            myConfigDataBL.SetModelFile();
        }

        private void BtnRunmodel_Click(object sender, RibbonControlEventArgs e)
        {
            ModelManagement myModel = new ModelManagement();
            myModel.RunWithAnimation(false);
        }


        private void BtnRunWindowless_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Not Implemented Yet");
        }

    }
}

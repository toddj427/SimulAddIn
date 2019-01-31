using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SimulationAddIn;

namespace SimulationAddIn.ShuttleWizard
{
    public partial class ShuttleWizard2 : Form
    {
        ConfigSheet myConfigSheet = new ConfigSheet();
        ShuttleDataBL myShuttleData = null;

        public ShuttleWizard2()
        {
            InitializeComponent();
        }

        public ShuttleWizard2(ShuttleDataBL argShuttleData) : this()
        {
            myShuttleData = argShuttleData;

        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
//            myShuttleData.SetArcFolder();
            if (Globals.ThisAddIn.WorksheetExists(ConfigRep.ConfigSheetName))
            {
                myConfigSheet.AddShuttleConfigs();
                MessageBox.Show("About to delete the Configuration sheet");
                Globals.ThisAddIn.DeleteSheetByName(ConfigRep.ConfigSheetName);
            }
            else
            {
                string title = "Warning";
                string message = "A configuration sheet does not yet exist";

                MessageBoxButtons buttons = MessageBoxButtons.OK;
                DialogResult result = MessageBox.Show(message, title, buttons);

            }

            ModelManagement myModel = new ModelManagement();
            myModel.CreateAMOFile();
            myModel.CreateASYFile();


        }
    }
}

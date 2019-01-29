using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimulationAddIn
{
    public class ShuttleDataBL
    {

        public ShuttleDataBL()
        {
        }

        // TODO: Need to change this to set a correct arc folder using code that works.
        public void SetArcFolder()
        {
            FolderBrowserDialog fD;

            fD = new FolderBrowserDialog();
            fD.Description = "Select the folder that contains or will contain the simulation model";
            fD.SelectedPath = ConfigRep.ReadModelPathFromSheet();
            DialogResult result = fD.ShowDialog();

            if (result == DialogResult.OK)
            {
                ConfigRep.WriteModelPathToSheet(fD.SelectedPath);
            }
        }
    }
}

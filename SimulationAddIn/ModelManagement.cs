using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace SimulationAddIn
{
    class ModelManagement
    {

        public void RunWithAnimation(bool syncMode)
        {
            string modelPath = ConfigRep.ReadModelPathFromSheet();
            string modelName = ConfigRep.ReadAutoModFileFromSheet();
            string fullPath = modelPath + "\\" + modelName + ".exe";

            MessageBox.Show(modelPath + "\\" + modelName + ".exe" );

            Process p = new Process();

            p.StartInfo.FileName = fullPath;
            p.EnableRaisingEvents = true;
            p.Exited += new EventHandler(RunHasExited);
            p.Start();
        }

        private void RunHasExited(object sender, System.EventArgs e)
        {
            MessageBox.Show("Simulation has exited");
        }
    }
}

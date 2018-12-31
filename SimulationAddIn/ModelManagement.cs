using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimulationAddIn
{
    class ModelManagement
    {

        public void RunWithAnimation(bool syncMode)
        {
            string modelPath = ConfigRep.ReadModelPathFromSheet();
            string modelName = ConfigRep.ReadAutoModFileFromSheet();

            MessageBox.Show(modelPath + "\\" + modelName + ".exe" );
        }
    }
}

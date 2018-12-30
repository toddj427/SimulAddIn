using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimulationAddIn
{
    class ConfigDataBL
    {
        public ConfigData configData { get; private set; }
        public ConfigDataBL()
        {
            configData = new ConfigData();
        }

        public void SetModelPath()
        {
            FolderBrowserDialog fD;

            fD = new FolderBrowserDialog();
            fD.Description = "Select the folder that contains or will contain the simulation model";
            DialogResult result  = fD.ShowDialog();

            if (result == DialogResult.OK)
            {
                configData.ModelPath = fD.SelectedPath;
            }
        }

        public string GetModelPath()
        {
            return configData.ModelPath;
        }

    }


}

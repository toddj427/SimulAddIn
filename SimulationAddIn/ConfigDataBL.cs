using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimulationAddIn
{
    // Used to manage the logic of the interface
    class ConfigDataBL
    {
        public ConfigDataBL()
        {
        }

        // Get the model path from the user and store it
        public void SetModelPath()
        {
            FolderBrowserDialog fD;

            fD = new FolderBrowserDialog();
            fD.Description = "Select the folder that contains or will contain the simulation model";
            fD.SelectedPath = ConfigRep.ReadModelPathFromSheet();
            DialogResult result  = fD.ShowDialog();

            if (result == DialogResult.OK)
            {
                ConfigRep.WriteModelPathToSheet(fD.SelectedPath);
            }
        }

        // Get the model name from the user and store it
        public void SetModelFile()
        {
            OpenFileDialog fD;

            fD = new OpenFileDialog();
            fD.Title = "Select the model executable file";
            fD.Filter = "Automod Executables|*.exe";
            fD.InitialDirectory = ConfigRep.ReadModelPathFromSheet();
            DialogResult result = fD.ShowDialog();

            if (result == DialogResult.OK)
            {

                string modelName = fD.FileName.Substring(0, fD.FileName.Length - 4);
                while (modelName.IndexOf('\\')>0)
                {
                    modelName = modelName.Substring(modelName.IndexOf('\\') + 1);
                }
                ConfigRep.WriteAutoModFileToSheet(modelName);
            }
        }
    }
}

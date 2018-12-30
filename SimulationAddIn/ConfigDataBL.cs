﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimulationAddIn
{
    class ConfigDataBL
    {
        public ConfigDataBL()
        {
        }

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

    }


}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Office = Microsoft.Office.Core;


namespace SimulationAddIn
{
    static class ConfigRep
    {
        public const string ConfigSheetName = "Config";

        static public void WriteModelPathToSheet(string myModelPath)
        {
            Excel.Worksheet myConfigSheet = Globals.ThisAddIn.GetWorkSheetByName(ConfigSheetName);

            myConfigSheet.Range["ModelPath"].Value = myModelPath;

        }
    }
}

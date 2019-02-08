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
    // Manage the storage and retrieval of the data
    // It is a static class because it does not persist any of the data it is reading or writing
    static class ConfigRep
    {
        public const string ConfigSheetName = "Config";

        #region ModelPath
        static public void WriteModelPathToSheet(string myModelPath)
        {
            Excel.Worksheet myConfigSheet = Globals.ThisAddIn.GetWorkSheetByName(ConfigSheetName);

            myConfigSheet.Range["ModelPath"].Value = myModelPath;
        }

        internal static string ReadModelPathFromSheet()
        {
            Excel.Worksheet myConfigSheet = Globals.ThisAddIn.GetWorkSheetByName(ConfigSheetName);

            return myConfigSheet.Range["ModelPath"].Value;
        }
        #endregion

        #region ModelName
        internal static void WriteAutoModFileToSheet(string fileName)
        {
            Excel.Worksheet myConfigSheet = Globals.ThisAddIn.GetWorkSheetByName(ConfigSheetName);

            myConfigSheet.Range["ModelName"].Value = fileName;
        }

        internal static string ReadAutoModFileFromSheet()
        {
            Excel.Worksheet myConfigSheet = Globals.ThisAddIn.GetWorkSheetByName(ConfigSheetName);

            return myConfigSheet.Range["ModelName"].Value;
        }
        #endregion  
    }
}

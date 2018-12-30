using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace SimulationAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public Excel.Worksheet GetActiveWorkSheet()
        {
            return (Excel.Worksheet)Application.ActiveSheet;
        }

        public Excel.Worksheet GetWorkSheetByName(string sheetName)
        {
            if (!WorksheetExists(sheetName))
            {
                return null;
            }
            else
            {
                return (Excel.Worksheet)Application.ActiveWorkbook.Worksheets[sheetName];
            }
        }

        public bool WorksheetExists(string sheetName)
        {
            Excel.Worksheet worksheet = null;
            // ToDo: Add the ability to see if a worksheet exists by name without using a try/catch
            try
            {
                worksheet = (Excel.Worksheet)Application.ActiveWorkbook.Worksheets[sheetName];
            }
            catch (Exception)
            {
                return false;
            }
            if (worksheet is null)
            {
                return false;
            }

            return true;
        }

        public Excel.Worksheet CreateNewSheet(string sheetName)
        {
            Excel.Worksheet newSheet = (Excel.Worksheet)Application.ActiveWorkbook.Sheets.Add();
            newSheet.Name = sheetName;
            return newSheet;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

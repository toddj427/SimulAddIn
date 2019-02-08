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
//            this.Application.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(Application_SheetActivate);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        // The below method, along with the line in the startup adds the ability for the add-in to detect when the active sheet is changed.
/*
        void Application_SheetActivate(object sh)
        {
            System.Windows.Forms.MessageBox.Show("Sheet is activated");
        }
*/

        // Return the active worksheet
        public Excel.Worksheet GetActiveWorkSheet()
        {
            return (Excel.Worksheet)Application.ActiveSheet;
        }

        // Get a worksheet pointer to the named worksheet 
        // returns null if the worksheet does not exist
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

        // This will remove the sheet named if it exists
        internal void DeleteSheetByName(string sheetName)
        {
            if (WorksheetExists(sheetName))
            {
                Excel.Worksheet worksheet = (Excel.Worksheet)Application.ActiveWorkbook.Worksheets[sheetName];
                worksheet.Delete();
            }
            return;
        }

        // Check to see if the worksheet exists
        public bool WorksheetExists(string sheetName)
        {
            Excel.Worksheet worksheet = null;
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

        // Create a new sheet in the active workbook and give it the name passed in
        public Excel.Worksheet CreateNewSheet(string sheetName)
        {
            Excel.Worksheet newSheet = (Excel.Worksheet)Application.ActiveWorkbook.Sheets.Add();
            newSheet.Name = sheetName;
            return newSheet;
        }

        // Return a list of all worksheets for the active workbook
        public Excel.Sheets GetWorkSheets()
        {
            Excel.Sheets theSheets = Application.ActiveWorkbook.Worksheets;
            return theSheets;
        }

        // Active the sheet that is named
        public void GoToSheet(string sheetName)
        {
            if (WorksheetExists(sheetName))
            {
                Application.ActiveWorkbook.Sheets[sheetName].Select();
            }
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

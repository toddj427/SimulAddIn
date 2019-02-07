using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;

namespace SimulationAddIn
{
    public partial class DataViewer : Form
    {
        public DataViewer()
        {
            InitializeComponent();
        }

        // Note that building up the list of categories and sheet types assumes that
        // There is a space on either side of a -
        // For example:   Category - SheetName
        // or: OMA Input - Input One
        private void DataViewer_Load(object sender, EventArgs e)
        {
            Excel.Sheets theSheets = Globals.ThisAddIn.GetWorkSheets();
            foreach (Excel.Worksheet theSheet in theSheets)
            {
                if (theSheet.Name.Contains("-"))
                {
                    string temps = theSheet.Name.Substring(0, theSheet.Name.IndexOf("-") - 1);
                    if (this.cbxFilter == null || !(this.cbxFilter.Items.Contains(temps)))
                    {
                        this.cbxFilter.Items.Add(temps);
                    }
                }
            }

        }

        private void cbxFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbxSheetList.Items.Clear();
            Excel.Sheets theSheets = Globals.ThisAddIn.GetWorkSheets();
            foreach (Excel.Worksheet theSheet in theSheets)
            {
                if (theSheet.Name.Contains("-"))
                {
                    string temps = theSheet.Name.Substring(theSheet.Name.IndexOf("-") + 2);
                    string tempType = theSheet.Name.Substring(0,theSheet.Name.IndexOf("-") - 1);
                    if (string.Equals(tempType,cbxFilter.SelectedItem.ToString()) && (this.lbxSheetList == null || !(this.lbxSheetList.Items.Contains(temps))))
                    {
                        this.lbxSheetList.Items.Add(temps);
                    }
                }
            }

        }

        private void lbxSheetList_SelectedIndexChanged(object sender, EventArgs e)
        {
            string temps = cbxFilter.SelectedItem.ToString() + " - " + lbxSheetList.SelectedItem.ToString();
            Globals.ThisAddIn.GoToSheet(temps);

        }
    }
}

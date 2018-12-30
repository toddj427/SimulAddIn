using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;


namespace SimulationAddIn
{
    // This class handles the formatting of the configuration sheet, including
    // colors, fonts, cell and sheet protection and names
    // data is handled by the config rep class
    class ConfigSheet
    {

        Excel.Worksheet configSheet = null;

        public ConfigSheet()
        {

        }


        public void InitializeNewSheet()
        {
            if (Globals.ThisAddIn.WorksheetExists(ConfigRep.ConfigSheetName))
            {
                MessageBox.Show("About to delete the Configuration sheet");
                Globals.ThisAddIn.DeleteSheetByName(ConfigRep.ConfigSheetName);
            }
            configSheet = Globals.ThisAddIn.CreateNewSheet(ConfigRep.ConfigSheetName);

            // Set up the colors
            configSheet.Range["B2:E20"].Borders[XlBordersIndex.xlEdgeLeft].Color = 0x175108;
            configSheet.Range["B2:E20"].Borders[XlBordersIndex.xlEdgeRight].Color = 0x175108;
            configSheet.Range["B2:E20"].Borders[XlBordersIndex.xlEdgeTop].Color = 0x175108;
            configSheet.Range["B2:E20"].Borders[XlBordersIndex.xlEdgeBottom].Color = 0x175108;
            configSheet.Range["B2:E20"].Interior.Color = 0xAEE2A1;

            // Set the configuration area title
            configSheet.Range["C3"].Value = "Simulation Configuration";
            configSheet.Range["C3"].Font.Size = 24;

            // Set the column widths for the key values
            configSheet.Columns[3].columnWidth = 18;
            configSheet.Columns[4].columnWidth = 26;

            // Add the model path cell
            configSheet.Range["C5"].Value = "Model Path: ";
            configSheet.Range["C5"].Cells.HorizontalAlignment = XlHAlign.xlHAlignRight;
            configSheet.Range["D5"].Interior.Color = 0xFFFFFF;
            configSheet.Range["D5"].Cells.Name = "ModelPath";

        }

       
    }
}

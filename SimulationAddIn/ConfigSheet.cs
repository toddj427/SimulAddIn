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
            if (Globals.ThisAddIn.WorksheetExists(ConfigRep.ConfigSheetName))
            {
                configSheet = Globals.ThisAddIn.GetWorkSheetByName(ConfigRep.ConfigSheetName);
            }
        }

        // Initialize the configuration sheet
        public void CreateNewConfigSheet()
        {
            if (Globals.ThisAddIn.WorksheetExists(ConfigRep.ConfigSheetName))
            {
                MessageBox.Show("About to delete the Configuration sheet");
                Globals.ThisAddIn.DeleteSheetByName(ConfigRep.ConfigSheetName);
            }
            configSheet = Globals.ThisAddIn.CreateNewSheet(ConfigRep.ConfigSheetName);

            // Set up the colors to use the Roar colors
            configSheet.Range["B2:E20"].Borders[XlBordersIndex.xlEdgeLeft].Color = System.Drawing.Color.FromArgb(0, 42, 94);
            configSheet.Range["B2:E20"].Borders[XlBordersIndex.xlEdgeRight].Color = System.Drawing.Color.FromArgb(0, 42, 94);
            configSheet.Range["B2:E20"].Borders[XlBordersIndex.xlEdgeTop].Color = System.Drawing.Color.FromArgb(0, 42, 94);
            configSheet.Range["B2:E20"].Borders[XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.FromArgb(0, 42, 94);
            configSheet.Range["B2:E20"].Interior.Color = System.Drawing.Color.FromArgb(0, 172, 228);

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

            // Add the model name cell
            configSheet.Range["C6"].Value = "Model Name: ";
            configSheet.Range["C6"].Cells.HorizontalAlignment = XlHAlign.xlHAlignRight;
            configSheet.Range["D6"].Interior.Color = 0xFFFFFF;
            configSheet.Range["D6"].Cells.Name = "ModelName";

        }

        public void AddShuttleConfigs()
        {
            // Add the number of Aisles
            configSheet.Range["C8"].Value = "Number of Aisles: ";
            configSheet.Range["C8"].Cells.HorizontalAlignment = XlHAlign.xlHAlignRight;
            configSheet.Range["D8"].Interior.Color = 0xFFFFFF;
            configSheet.Range["D8"].Cells.Name = "NumAisles";
            configSheet.Range["D8"].Cells.Value = 3;

            // Add the number of Bays
            configSheet.Range["C9"].Value = "Number of Bays: ";
            configSheet.Range["C9"].Cells.HorizontalAlignment = XlHAlign.xlHAlignRight;
            configSheet.Range["D9"].Interior.Color = 0xFFFFFF;
            configSheet.Range["D9"].Cells.Name = "NumBays";
            configSheet.Range["D9"].Cells.Value = 5;

            // Add the number of Shelves
            configSheet.Range["C10"].Value = "Number of Shelves: ";
            configSheet.Range["C10"].Cells.HorizontalAlignment = XlHAlign.xlHAlignRight;
            configSheet.Range["D10"].Interior.Color = 0xFFFFFF;
            configSheet.Range["D10"].Cells.Name = "NumShelves";
            configSheet.Range["D10"].Cells.Value = 3;

        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace SimulationAddIn
{
    public partial class ThisAddIn
    {
        private SafeNativeMethods.HookProc _keyboardProc;
        private IntPtr _hookIdKeyboard;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _keyboardProc = KeyboardHookCallback;
            SetWindowsHooks();
           
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

        /******************************/
        /* Keyboard hook code */
        /******************************/
        private void SetWindowsHooks()
        {
            uint threadId = (uint)SafeNativeMethods.GetCurrentThreadId();

            _hookIdKeyboard =
                SafeNativeMethods.SetWindowsHookEx(
                    (int)SafeNativeMethods.WH_KEYBOARD,
                    _keyboardProc,
                    IntPtr.Zero,
                    threadId);
        }

        private void UnhookWindowsHooks()
        {
            SafeNativeMethods.UnhookWindowsHookEx(_hookIdKeyboard);
        }


        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public struct KBDLLHookStruct
        {
            public Int32 vkCode;
            public Int32 scanCode;
            public Int32 flags;
            public Int32 time;
            public Int32 dwExtraInfo;
        }

        public enum KeyEvents
        {
            KeyDown = 0x0100,
            KeyUp = 0x0101,
            SKeyDown = 0x0104,
            SKeyUp = 0x0105
        }

        static bool wasControl = false;
        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
            {
                return SafeNativeMethods.CallNextHookEx(_hookIdKeyboard, nCode, wParam, lParam);
            }

            if (nCode == 3)
            {
                IntPtr ptr = IntPtr.Zero;

                int keydownup = (int)lParam >> 30;
                if (keydownup == 0)
                {
                    if ((int)wParam == 0x11)
                    {
                        wasControl = true;
                    }
                    if (wasControl == true && (int)wParam == 0x4d)
                    {
                        System.Windows.Forms.MessageBox.Show("Control m was pushed");
                        wasControl = false;

                    }
                    // It is a down key
                }
                if (keydownup == -1)
                {
                    wasControl = false;
                }
            }
            return SafeNativeMethods.CallNextHookEx(_hookIdKeyboard,nCode,wParam,lParam);
        }

        // Return the active worksheet
        public Excel.Worksheet GetActiveWorkSheet()
        {
            return (Excel.Worksheet)Application.ActiveSheet;
        }
        /******************************/
        /* end of */
        /******************************/

        #region worksheetFuncs
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
        #endregion

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

    internal static class SafeNativeMethods
    {
        public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        public static int WH_KEYBOARD = 2;

        public enum WindowMessages : uint
        {
            WM_KEYDOWN = 0x0100,
            WM_KEYFIRST = 0x0100,
            WM_KEYLAST = 0x0108,
            WM_KEYUP = 0x0101,
            WM_LBUTTONDBLCLK = 0x0203,
            WM_LBUTTONDOWN = 0x0201,
            WM_LBUTTONUP = 0x0202,
            WM_MBUTTONDBLCLK = 0x0209,
            WM_MBUTTONDOWN = 0x0207,
            WM_MBUTTONUP = 0x0208,
            WM_MOUSEACTIVATE = 0x0021,
            WM_MOUSEFIRST = 0x0200,
            WM_MOUSEHOVER = 0x02A1,
            WM_MOUSELAST = 0x020D,
            WM_MOUSELEAVE = 0x02A3,
            WM_MOUSEMOVE = 0x0200,
            WM_MOUSEWHEEL = 0x020A,
            WM_MOUSEHWHEEL = 0x020E,
            WM_RBUTTONDBLCLK = 0x0206,
            WM_RBUTTONDOWN = 0x0204,
            WM_RBUTTONUP = 0x0205,
            WM_SYSDEADCHAR = 0x0107,
            WM_SYSKEYDOWN = 0x0104,
            WM_SYSKEYUP = 0x0105
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr CallNextHookEx(
            IntPtr hhk,
            int nCode,
            IntPtr wParam,
            IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetCurrentThreadId();
    }

}

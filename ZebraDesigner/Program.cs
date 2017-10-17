using OfficeOpenXml;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace ZebraDesigner
{
    class Program
    {
        //include FindWindow
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        //include FindWindowEx
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        //include SendMessage
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, IntPtr lParam);

        //include Mouse click events
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, int dx, int dy, uint cButtons, IntPtr dwExtraInfo);

        //include Mouse cursor location
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);

        //Main programm to look at label file, excel file and row count.
        //**********************************************************************************************//
        //*********************************************************************************************//
        /*****/          static void Main()                                                     /******/
        /*****/              {                                                                 /******/                          
        /*****/                  Zebra("Plauktiem dzeltena.lbl", "Book1.xlsx", 1001);         /******/                           
        /*****/              }                                                               /******/                            
        /*****/                                                                             /******/               
        //****************************************************************************************/
        //***************************************************************************************/
        //static void Zebra(string File, string Amount)
        static void Zebra(string File, string ExcFile, int daudzums)
        {
            // OPEN LABEL //
            Process Label = Process.Start(@"C:\temp\" + File);
            Thread.Sleep(7000);

            // SEARCH HANDLES & CHECK OK FOR PRINTERS //
            const int BN_CLICKED = 245;
            IntPtr ZebraDesigner = FindWindow(null, "Design");
            IntPtr ChoosePrinterDialog = FindWindowEx(ZebraDesigner, IntPtr.Zero, "#32770", "Select Printer");
            IntPtr ButtonOK = FindWindowEx(ChoosePrinterDialog, IntPtr.Zero, "Button", "OK");
            SendMessage((IntPtr)ButtonOK, BN_CLICKED, 0, IntPtr.Zero);
            Thread.Sleep(1000);



            // LOOP THROUGHT EXCEL FILE


            string curr = Directory.GetCurrentDirectory();
            FileInfo fileName = new FileInfo(@"C:\temp\"+ ExcFile);
            ExcelPackage pck = new ExcelPackage(fileName);
            var ws = pck.Workbook.Worksheets["Sheet1"];
            var row = 2;
            var acol = 2;
            var bcol = 3;

            for (row = 2; row <= daudzums; row++)
            {
                if (ws.Cells[row, acol].Value == null)
                {
                    goto end;
                }


                //Double mouse left click on program center
                Cursor.Position = new System.Drawing.Point(800, 550);
                SendMouseDoubleClick();
                Thread.Sleep(1000);

                // edit text

                IntPtr TextWizzard = FindWindowEx(ZebraDesigner, IntPtr.Zero, "#32770", "Text Wizard");
                //IntPtr Cont = FindWindowEx(TextWizzard, IntPtr.Zero, "Button", "&Content");
                //IntPtr Text = FindWindowEx(TextWizzard, IntPtr.Zero, "Edit", "I002-2");

                //Message(Text, 0x000C, IntPtr.Zero, new StringBuilder("Hello World!"));
                Thread.Sleep(1000);
                for (int i = 1; i <= 10; i++)
                {
                    SendKeys.SendWait("{DEL}");
                }
                Thread.Sleep(1000);
                SendKeys.SendWait(ws.Cells[row, acol].Value.ToString());
                Thread.Sleep(500);
                IntPtr ButtonFinish = FindWindowEx(TextWizzard, IntPtr.Zero, "Button", "&Finish");
                SendMessage((IntPtr)ButtonFinish, BN_CLICKED, 0, IntPtr.Zero);

                Thread.Sleep(500);
                Cursor.Position = new System.Drawing.Point(1400, 550);
                SendMouseDoubleClick();

                Thread.Sleep(1000);
                for (int i = 1; i <= 5; i++)
                {
                    SendKeys.SendWait("{DEL}");
                }
                Thread.Sleep(1000);
                SendKeys.SendWait(ws.Cells[row, bcol].Value.ToString());
                Thread.Sleep(500);
                IntPtr BarCodeWizard = FindWindowEx(ZebraDesigner, IntPtr.Zero, "#32770", "Bar Code Wizard");
                IntPtr ButtonFinish2 = FindWindowEx(BarCodeWizard, IntPtr.Zero, "Button", "&Finish");
                SendMessage((IntPtr)ButtonFinish2, BN_CLICKED, 0, IntPtr.Zero);


                // SHORTCUT TO PRINTING //
                SendKeys.SendWait("^(p)");
                Thread.Sleep(1000);

                // EDIT AMOUNT //
                IntPtr PrintDialog = FindWindowEx(ZebraDesigner, IntPtr.Zero, "#32770", "Print");
                IntPtr ChnageAmount = FindWindowEx(PrintDialog, IntPtr.Zero, "Edit", null);
                SendKeys.SendWait("1");
                Thread.Sleep(1000);

                // CHECKBOXES //
                const int BM_SETCHECK = 0x00f1;
                const int BST_CHECKED = 0x0001;
                const int BST_UNCHECKED = 0x0000;
                IntPtr PrintToFile = FindWindowEx(PrintDialog, IntPtr.Zero, "Button", "Print to &file");
                IntPtr CloseAfterPrint = FindWindowEx(PrintDialog, IntPtr.Zero, "Button", "Close after &print");
                SendMessage(PrintToFile, BM_SETCHECK, BST_UNCHECKED, IntPtr.Zero);
                SendMessage(CloseAfterPrint, BM_SETCHECK, BST_CHECKED, IntPtr.Zero);
                Thread.Sleep(1000);

                // PRINT //
                //     IntPtr Print = FindWindowEx(PrintDialog, IntPtr.Zero, "Button", "Print");
                IntPtr Print = FindWindowEx(PrintDialog, IntPtr.Zero, "Button", "Print");
                SendMessage((IntPtr)Print, BN_CLICKED, 0, IntPtr.Zero);
                Thread.Sleep(2000);

            }



        // COSE WINDOWS //
        end:
            Thread.Sleep(1000);
            SendKeys.SendWait("%{F4}");
            Thread.Sleep(1000);
            IntPtr SaveFileDialog = FindWindowEx(ZebraDesigner, IntPtr.Zero, "#32770", "ZebraDesigner - Plauktiem dzeltena.lbl");
            IntPtr ButtonNo = FindWindowEx(SaveFileDialog, IntPtr.Zero, "Button", "&No");
            SendMessage((IntPtr)ButtonNo, BN_CLICKED, 0, IntPtr.Zero);
            Thread.Sleep(1000);
        }

        private const uint MOUSEEVENTF_LEFTDOWN = 0x02;
        private const uint MOUSEEVENTF_LEFTUP = 0x04;
        public static void SendMouseDoubleClick()
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);

            Thread.Sleep(150);

            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero);
        }
                
    }
}

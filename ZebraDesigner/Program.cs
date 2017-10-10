using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
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
        public static extern int SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int msg, int wParam, StringBuilder lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, int dx, int dy, uint cButtons, uint dwExtraInfo);




        static void Main()
        {
            Zebra("113-156-1000 BY.lbl");
        }


        //static void Zebra(string File, string Amount)
        static void Zebra(string File)
        {
            // OPEN LABEL //
            Process Label = Process.Start(@"C:\Users\Oskars.Vanags\Desktop\JYSK\random stuff\LABEL\" + File);
            System.Threading.Thread.Sleep(7000);

            // SEARCH HANDLES & CHECK OK FOR PRINTERS //
            const int BN_CLICKED = 245;
            IntPtr ZebraDesigner = FindWindow(null, "Design");
            IntPtr ChoosePrinterDialog = FindWindowEx(ZebraDesigner, IntPtr.Zero, "#32770", "Select Printer");
            IntPtr ButtonOK = FindWindowEx(ChoosePrinterDialog, IntPtr.Zero, "Button", "OK");
            SendMessage((IntPtr)ButtonOK, BN_CLICKED, 0, 0);
            System.Threading.Thread.Sleep(1000);

            const int MOUSEEVENTF_LEFTDOWN = 0x02;
            const int MOUSEEVENTF_LEFTUP = 0x04;
            int X = Cursor.Position.X;
            int Y = Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_LEFTDOWN, X, Y, 0, 0);
            System.Threading.Thread.Sleep(1);
            mouse_event(MOUSEEVENTF_LEFTUP, X, Y, 0, 0);
            System.Threading.Thread.Sleep(1);
            mouse_event(MOUSEEVENTF_LEFTDOWN, X, Y, 0, 0);
            System.Threading.Thread.Sleep(1);
            mouse_event(MOUSEEVENTF_LEFTUP, X, Y, 0, 0);
            System.Threading.Thread.Sleep(1000);

            const int WM_GETTEXTLENGTH = 0x000E;
            const int WM_GETTEXT = 0x000D;
            IntPtr TextWizzard = FindWindowEx(ZebraDesigner, IntPtr.Zero, "#32770", "Text Wizard");
            IntPtr Cont = FindWindowEx(TextWizzard, IntPtr.Zero, "Button", "&Content");
            IntPtr Text = FindWindowEx(Cont, IntPtr.Zero, "Edit", null);
            System.Threading.Thread.Sleep(500);
            int length = SendMessage(Text, WM_GETTEXTLENGTH, 0, 0);
            StringBuilder text3 = new StringBuilder(length);
            int hr = SendMessage(Text, WM_GETTEXT, length, text3);
            Console.WriteLine(length);
            Console.WriteLine(text3);
            Console.ReadKey();

            /*
                        // SHORTCUT TO PRINTING //
                        SendKeys.SendWait("^(p)");
                        System.Threading.Thread.Sleep(1000);

                        // EDIT AMOUNT //
                        const int WM_SETTEXT = 0x000c;
                        IntPtr PrintDialog = FindWindowEx(ZebraDesigner, IntPtr.Zero, "#32770", "Print");
                        IntPtr ChnageAmount = FindWindowEx(PrintDialog, IntPtr.Zero, "Edit", null);
                        SendMessage((int)ChnageAmount, WM_SETTEXT, 0, Amount);
                        System.Threading.Thread.Sleep(1000);

                        // CHECKBOXES //
                        const int BM_SETCHECK = 0x00f1;
                        const int BST_CHECKED = 0x0001;
                        const int BST_UNCHECKED = 0x0000;
                        IntPtr PrintToFile = FindWindowEx(PrintDialog, IntPtr.Zero, "Button", "Print to &file");
                        IntPtr CloseAfterPrint = FindWindowEx(PrintDialog, IntPtr.Zero, "Button", "Close after &print");
                        SendMessage((int)PrintToFile, BM_SETCHECK, BST_CHECKED, null);
                        SendMessage((int)CloseAfterPrint, BM_SETCHECK, BST_UNCHECKED, null);
                        System.Threading.Thread.Sleep(1000);

                        // PRINT //
                        IntPtr Print = FindWindowEx(PrintDialog, IntPtr.Zero, "Button", "Print");

                        // COSE WINDOWS //
                        */


        }
    }
}

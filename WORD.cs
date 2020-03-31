using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace TENKA_ÖĞRENCİ_PANELİ
{
    public partial class WORD : Form
    {
        public WORD()
        {
            InitializeComponent();
        }

        [DllImport("user32.dll")]
        static extern IntPtr SetParent(IntPtr child, IntPtr newParent);
        [DllImport("user32.dll")]
        static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int IParam);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        private const int WM_SYSCOMMAND = 274;
        private const int SC_MAXIMIZE = 61488;
        private void WORD_Load(object sender, EventArgs e)
        {
            string exeyolu = "cmd.exe"; 
            Process calistir = Process.Start(exeyolu);
            while (calistir.MainWindowHandle == IntPtr.Zero || IsWindowVisible(calistir.MainWindowHandle))
            {
                System.Threading.Thread.Sleep(10);
                calistir.Refresh();
            }
            
            SetParent(calistir.MainWindowHandle, this.Handle);
            _ = SendMessage(calistir.MainWindowHandle, WM_SYSCOMMAND, SC_MAXIMIZE, 0);
        }
    }
}

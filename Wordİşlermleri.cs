using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace TENKA_ÖĞRENCİ_PANELİ
{
    public partial class Wordİşlermleri : Form
    {
        public Wordİşlermleri()
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
        private void bunifuFlatButton2_Click(object sender, EventArgs e)
        {
            
            Application wordApp = new Application();
            object missing = System.Reflection.Missing.Value;
            Document doc = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            wordApp.Visible = true;
            doc.Content.SetRange(0, 0);
            doc.Content.Text = @"C:\Users\MUHAMMED ACIBALIK\Masaüstü";
            object fileName = @"C:\Users\MUHAMMED ACIBALIK\Masaüstü";
            doc.SaveAs2(ref fileName);
            doc.Close(ref missing, ref missing, ref missing);
            doc = null;
            wordApp.Quit(ref missing, ref missing, ref missing);
            wordApp = null;
            System.Windows.Forms.MessageBox.Show("Word dosyası başarılı bir şekilde oluşturuldu.!");
        }

        private void bunifuFlatButton1_Click(object sender, EventArgs e)
        {

        }

        private void bunifuFlatButton1_Click_1(object sender, EventArgs e)
        {
            WORD wORD = new WORD();
            wORD.Show();
        }

        private void bunifuFlatButton2_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Word Dosyası |*.docx";
            if (openFileDialog1.ShowDialog()==DialogResult.OK)
            {
                bunifuFlatButton1.Name = openFileDialog1.FileName;
                
            }
            System.Diagnostics.Process.Start(openFileDialog1.FileName);
        }
    }
}



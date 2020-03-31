using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BunifuLC;

namespace TENKA_ÖĞRENCİ_PANELİ
{
    public partial class Çalışmalarım : Form
    {
        public Çalışmalarım()
        {
            InitializeComponent();
        }
        string dizin = string.Empty;
        string dosyaAdi = string.Empty;
        private void bunifuTileButton1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton1.LabelText = openFileDialog1.FileName;
                dizin = bunifuTileButton1.LabelText = openFileDialog1.SafeFileName;
            }

        }

        private void bunifuThinButton21_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog1.FileName);
        }

        private void bunifuTileButton3_Click(object sender, EventArgs e)
        {
            openFileDialog3.Filter = "All files (*.*)|*.*";
            if (openFileDialog3.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
               dosyaAdi = bunifuTileButton3.LabelText = openFileDialog3.FileName;
               dizin = bunifuTileButton3.LabelText = openFileDialog3.SafeFileName;
            }
        }

        private void bunifuTileButton5_Click(object sender, EventArgs e)
        {
            openFileDialog5.Filter = "All files (*.*)|*.*";
            if (openFileDialog5.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton5.LabelText = openFileDialog5.FileName;
                dizin = bunifuTileButton5.LabelText = openFileDialog5.SafeFileName;
            }
        }

        private void bunifuTileButton7_Click(object sender, EventArgs e)
        {
            openFileDialog7.Filter = "All files (*.*)|*.*";
            if (openFileDialog7.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
               dosyaAdi = bunifuTileButton7.LabelText = openFileDialog7.FileName;
               dizin = bunifuTileButton7.LabelText = openFileDialog7.SafeFileName;
            }
        }

        private void bunifuTileButton9_Click(object sender, EventArgs e)
        {
            openFileDialog9.Filter = "All files (*.*)|*.*";
            if (openFileDialog9.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton9.LabelText = openFileDialog9.FileName;
                dizin = bunifuTileButton9.LabelText = openFileDialog9.SafeFileName;
            }
        }

        private void bunifuTileButton11_Click(object sender, EventArgs e)
        {
            openFileDialog11.Filter = "All files (*.*)|*.*";
            if (openFileDialog11.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton11.LabelText = openFileDialog11.FileName;
                dizin = bunifuTileButton11.LabelText = openFileDialog11.SafeFileName;
            }
        }

        private void bunifuTileButton2_Click(object sender, EventArgs e)
        {
            openFileDialog2.Filter = "All files (*.*)|*.*";
            if (openFileDialog2.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton2.LabelText = openFileDialog2.FileName;
                dizin = bunifuTileButton2.LabelText = openFileDialog2.SafeFileName;
            }
        }

        private void bunifuTileButton4_Click(object sender, EventArgs e)
        {
            openFileDialog4.Filter = "All files (*.*)|*.*";
            if (openFileDialog4.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton4.LabelText = openFileDialog4.FileName;
                dizin = bunifuTileButton4.LabelText = openFileDialog4.SafeFileName;
            }
        }

        private void bunifuTileButton6_Click(object sender, EventArgs e)
        {
            openFileDialog6.Filter = "All files (*.*)|*.*";
            if (openFileDialog6.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton6.LabelText = openFileDialog6.FileName;
                dizin = bunifuTileButton6.LabelText = openFileDialog6.SafeFileName;
            }
        }

        private void bunifuTileButton8_Click(object sender, EventArgs e)
        {
            openFileDialog8.Filter = "All files (*.*)|*.*";
            if (openFileDialog8.ShowDialog() == DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton8.LabelText = openFileDialog8.FileName;
                dizin = bunifuTileButton8.LabelText = openFileDialog8.SafeFileName;
            }
        }

        private void bunifuTileButton10_Click(object sender, EventArgs e)
        {
            openFileDialog10.Filter = "All files (*.*)|*.*";
            if (openFileDialog10.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton10.LabelText = openFileDialog10.FileName;
                dizin = bunifuTileButton10.LabelText = openFileDialog10.SafeFileName;
            }
        }

        private void bunifuTileButton12_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "All files (*.*)|*.*";
            if (openFileDialog12.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton12.LabelText = openFileDialog12.FileName;
                dizin = bunifuTileButton12.LabelText = openFileDialog12.SafeFileName;
            }
        }

        private void bunifuTileButton13_Click(object sender, EventArgs e)
        {
            openFileDialog13.Filter = "All files (*.*)|*.*";
            if (openFileDialog13.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton13.LabelText = openFileDialog13.FileName;
                dizin = bunifuTileButton13.LabelText = openFileDialog13.SafeFileName;
            }
        }

        private void bunifuTileButton14_Click(object sender, EventArgs e)
        {
            openFileDialog14.Filter = "All files (*.*)|*.*";
            if (openFileDialog14.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton14.LabelText = openFileDialog14.FileName;
                dizin = bunifuTileButton14.LabelText = openFileDialog14.SafeFileName;
            }
        }

        private void bunifuTileButton15_Click(object sender, EventArgs e)
        {
            openFileDialog15.Filter = "All files (*.*)|*.*";
            if (openFileDialog15.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton15.LabelText = openFileDialog15.FileName;
                dizin = bunifuTileButton15.LabelText = openFileDialog15.SafeFileName;
            }
        }

        private void bunifuTileButton16_Click(object sender, EventArgs e)
        {
            openFileDialog16.Filter = "All files (*.*)|*.*";
            if (openFileDialog16.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton16.LabelText = openFileDialog16.FileName;
                dizin = bunifuTileButton16.LabelText = openFileDialog16.SafeFileName;
            }
        }

        private void bunifuTileButton17_Click(object sender, EventArgs e)
        {
            openFileDialog17.Filter = "All files (*.*)|*.*";
            if (openFileDialog17.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton17.LabelText = openFileDialog17.FileName;
                dizin = bunifuTileButton17.LabelText = openFileDialog17.SafeFileName;
            }
        }

        private void bunifuTileButton18_Click(object sender, EventArgs e)
        {
            openFileDialog18.Filter = "All files (*.*)|*.*";
            if (openFileDialog18.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton18.LabelText = openFileDialog18.FileName;
                dizin = bunifuTileButton18.LabelText = openFileDialog18.SafeFileName;
            }
        }

        private void bunifuTileButton19_Click(object sender, EventArgs e)
        {
            openFileDialog19.Filter = "All files (*.*)|*.*";
            if (openFileDialog19.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton19.LabelText = openFileDialog19.FileName;
                dizin = bunifuTileButton19.LabelText = openFileDialog19.SafeFileName;
            }
        }

        private void bunifuTileButton20_Click(object sender, EventArgs e)
        {
            openFileDialog20.Filter = "All files (*.*)|*.*";
            if (openFileDialog20.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton20.LabelText = openFileDialog20.FileName;
                dizin = bunifuTileButton20.LabelText = openFileDialog20.SafeFileName;
            }
        }

        private void bunifuTileButton21_Click(object sender, EventArgs e)
        {
            openFileDialog21.Filter = "All files (*.*)|*.*";
            if (openFileDialog21.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton21.LabelText = openFileDialog21.FileName;
                dizin = bunifuTileButton21.LabelText = openFileDialog21.SafeFileName;
            }
        }

        private void bunifuTileButton22_Click(object sender, EventArgs e)
        {
            openFileDialog22.Filter = "All files (*.*)|*.*";
            if (openFileDialog22.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton22.LabelText = openFileDialog22.FileName;
                dizin = bunifuTileButton22.LabelText = openFileDialog22.SafeFileName;
            }
        }

        private void bunifuTileButton23_Click(object sender, EventArgs e)
        {
            openFileDialog23.Filter = "All files (*.*)|*.*";
            if (openFileDialog23.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton23.LabelText = openFileDialog23.FileName;
                dizin = bunifuTileButton23.LabelText = openFileDialog23.SafeFileName;
            }
        }

        private void bunifuTileButton24_Click(object sender, EventArgs e)
        {
            openFileDialog24.Filter = "All files (*.*)|*.*";
            if (openFileDialog24.ShowDialog()==DialogResult.OK)
            {
                string dizin = string.Empty;
                string dosyaAdi = string.Empty;
                dosyaAdi = bunifuTileButton24.LabelText = openFileDialog24.FileName;
                dizin = bunifuTileButton24.LabelText = openFileDialog24.SafeFileName;
            }
        }

        private void bunifuThinButton22_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog2.FileName);
        }

        private void bunifuThinButton23_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog3.FileName);
        }

        private void bunifuThinButton24_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog4.FileName);
        }

        private void bunifuThinButton25_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog5.FileName);
        }

        private void bunifuThinButton26_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog6.FileName);
        }

        private void bunifuThinButton27_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog7.FileName);
        }

        private void bunifuThinButton28_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog8.FileName);
        }

        private void bunifuThinButton29_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog9.FileName);
        }

        private void bunifuThinButton210_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog10.FileName);
        }

        private void bunifuThinButton211_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog11.FileName);
        }

        private void bunifuThinButton212_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog12.FileName);
        }

        private void bunifuThinButton213_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog13.FileName);
        }

        private void bunifuThinButton214_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog14.FileName);
        }

        private void bunifuThinButton215_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog15.FileName);
        }

        private void bunifuThinButton216_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog16.FileName);
        }

        private void bunifuThinButton217_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog17.FileName);
        }

        private void bunifuThinButton218_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog18.FileName);
        }

        private void bunifuThinButton219_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog19.FileName);
        }

        private void bunifuThinButton220_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog20.FileName);
        }

        private void bunifuThinButton221_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog21.FileName);
        }

        private void bunifuThinButton222_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog22.FileName);
        }

        private void bunifuThinButton223_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog23.FileName);
        }

        private void bunifuThinButton224_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(openFileDialog24.FileName);
        }

        private void Çalışmalarım_Load(object sender, EventArgs e)
        {

        }
    }
}

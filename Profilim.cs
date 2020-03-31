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
using Bunifu.Framework;

namespace TENKA_ÖĞRENCİ_PANELİ
{
    public partial class Profilim : Form
    {
        int tıklama = 0;
        public Profilim()
        {
            InitializeComponent();
        }

        private void bunifuRadialGauge1_ValueChanged(object sender, Bunifu.UI.WinForms.BunifuRadialGauge.ValueChangedEventArgs e)
        {

        }

        private void bunifuTextBox1_TextChanged(object sender, EventArgs e)
        {
            
            if (bunifuTextBox1.Text==""&bunifuTextBox1.Text==string.Empty)
            {
                this.bunifuRadialGauge1.Value = 0;
            }
            else
            {
                this.bunifuRadialGauge1.Value = 20;
            }
        }

        private void bunifuTextBox2_TextChanged(object sender, EventArgs e)
        {

            
            if (bunifuTextBox2.Text==""&bunifuTextBox2.Text==string.Empty)
            {
                this.bunifuRadialGauge1.Value = 20;
            }
            else
            {
                this.bunifuRadialGauge1.Value = 40;
            }
        }

        private void bunifuTextBox3_TextChanged(object sender, EventArgs e)
        {
            if (bunifuTextBox3.Text==""&bunifuTextBox3.Text==string.Empty)
            {
                this.bunifuRadialGauge1.Value = 40;
            }
            
            this.bunifuRadialGauge1.Value = 60;

        }

        private void bunifuDatepicker1_onValueChanged(object sender, EventArgs e)
        {
           
            this.bunifuRadialGauge1.Value = 100;

        }

        private void bunifuTextBox4_TextChanged(object sender, EventArgs e)
        {
            if (bunifuTextBox4.Text==""&bunifuTextBox4.Text==string.Empty)
            {
                this.bunifuRadialGauge1.Value = 60;
            }
            else
            {
                this.bunifuRadialGauge1.Value = 80;
            }
        }

        private void bunifuTextBox8_TextChanged(object sender, EventArgs e)
        {
            string aa = bunifuTextBox8.Text;
            this.bunifuRadialGauge2.Value = 25 / 2;
        }

        private void bunifuTextBox7_TextChanged(object sender, EventArgs e)
        {
            string bb = bunifuTextBox7.Text;
            this.bunifuRadialGauge2.Value = 25;
        }

        private void bunifuTextBox6_TextChanged(object sender, EventArgs e)
        {
            string cc = bunifuTextBox6.Text;
            this.bunifuRadialGauge2.Value = 75 / 2;
        }

        private void bunifuTextBox5_TextChanged(object sender, EventArgs e)
        {
            string dd = bunifuTextBox5.Text;
            this.bunifuRadialGauge2.Value = 50;
        }

        private void bunifuTextBox9_TextChanged(object sender, EventArgs e)
        {
            string ee = bunifuTextBox9.Text;
            this.bunifuRadialGauge2.Value = 125 / 2;
        }
        private void bunifuDatepicker2_onValueChanged(object sender, EventArgs e)
        {
            if (bunifuRadioButton2.Checked==true)
            {
                this.bunifuRadialGauge2.Value = 175 / 2;
            }
            else
            {
                this.bunifuRadialGauge2.Value = 200 / 2;
            }

        }

        private void bunifuDatepicker3_onValueChanged(object sender, EventArgs e)
        {
            
            this.bunifuRadialGauge2.Value = 100;
        }

        private void bunifuButton1_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory(@"C:\ProgramData\Tenka\KişiselBilgi");
            Directory.CreateDirectory(@"C:\ProgramData\Tenka\EğitimBilgi");

            using (TextWriter a = new StreamWriter(@"C:\ProgramData\Tenka\KişiselBilgi\kişiselbilgi.text")) ;
            using (TextWriter b = new StreamWriter(@"C:\ProgramData\Tenka\EğitimBilgi\eğitimbilgi.text")) ;

            using (StreamWriter aa = new StreamWriter(@"C:\ProgramData\Tenka\KişiselBilgi\kişiselbilgi.text"))
            {
                aa.WriteLine(bunifuTextBox1.Text);
                aa.WriteLine(bunifuTextBox2.Text);
                aa.WriteLine(bunifuTextBox3.Text);
                aa.WriteLine(bunifuTextBox4.Text);
                aa.WriteLine(bunifuDatepicker1.Value);
                aa.WriteLine(bunifuDropdown1.SelectedItem);
            }
            using (StreamWriter bb = new StreamWriter(@"C:\ProgramData\Tenka\EğitimBilgi\eğitimbilgi.text"))
            {
                bb.WriteLine(bunifuTextBox8.Text);
                bb.WriteLine(bunifuTextBox7.Text);
                bb.WriteLine(bunifuTextBox6.Text);
                bb.WriteLine(bunifuTextBox5.Text);
                bb.WriteLine(bunifuTextBox9.Text);
                bb.WriteLine(bunifuLabel4.Text);
                bb.WriteLine(bunifuLabel5.Text);
                bb.WriteLine(bunifuDatepicker2.Value);
                bb.WriteLine(bunifuDatepicker3.Value);
            }
            
            bunifuButton1.Text = "GÜNCELLE";
            tıklama++;
            if (tıklama%2==1)
            {
                MessageBox.Show("BİLGİLERİNİZ BAŞARI İLE KAYDEDİLMİŞTİR");
            }         if (tıklama %2==0)
            {
                
                MessageBox.Show("BİLGİLERİNİZ BAŞARIYLA GÜNCELLENMİŞTİR");
            }

            bool ff = bunifuRadioButton1.Checked = true;
            this.bunifuRadialGauge2.Value = 175 / 2;
        }

        private void bunifuButton2_Click(object sender, EventArgs e)
        {
            string[] kb = File.ReadAllLines(@"C:\ProgramData\Tenka\KişiselBilgi\kişiselbilgi.text");
            if (kb.Length > 0)
                bunifuTextBox1.Text = kb[0];
            bunifuTextBox2.Text = kb[1];
            bunifuTextBox3.Text = kb[2];
            bunifuTextBox4.Text = kb[3];
            bunifuDropdown1.Text = kb[5];

            string[] eb = File.ReadAllLines(@"C:\ProgramData\Tenka\EğitimBilgi\eğitimbilgi.text");
            if (eb.Length > 0)
                bunifuTextBox8.Text = eb[0];
            bunifuTextBox7.Text = eb[1];
            bunifuTextBox6.Text = eb[2];
            bunifuTextBox5.Text = eb[3];
            bunifuTextBox9.Text = eb[4];
            if (eb[5]==bunifuLabel4.Text)
            {
                bunifuRadioButton1.Checked = true;
                bunifuDatepicker3.Visible = false;
                bunifuLabel3.Visible = false;
                bunifuRadialGauge2.Value = 175 / 2;
            }
            else
            {
                bunifuRadioButton2.Checked = true;
                bunifuRadialGauge2.Value = 149 / 2;
            }
        }
        
        private void Profilim_Load(object sender, EventArgs e)
        {
        }

        private void bunifuRadioButton1_Click(object sender, EventArgs e)
        {
            if (bunifuRadioButton1.Checked==true)
            {
                bunifuDatepicker3.Visible = false;
                bunifuLabel3.Visible = false;
            }
            bunifuRadialGauge2.Value = 81;

        }

        private void bunifuRadioButton2_Click(object sender, EventArgs e)
        {
            if (bunifuRadioButton2.Checked==true)
            {
                bunifuDatepicker3.Visible = true;
                bunifuLabel3.Visible = true;
            }
        }

        private void bunifuLabel4_Click(object sender, EventArgs e)
        {

        }
    }
}

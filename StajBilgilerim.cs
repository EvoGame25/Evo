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

namespace TENKA_ÖĞRENCİ_PANELİ
{
    public partial class StajBilgilerim : Form
    {
        public StajBilgilerim()
        {
            InitializeComponent();
        }

        private void bunifuFlatButton1_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory(@"C:\ProgramData\Tenka\Stajbilgileri");
            using(StreamWriter a = new StreamWriter(@"C:\ProgramData\Tenka\Stajbilgileri\KişiselBilgierim.text"))
            {
                a.WriteLine(bunifuTextBox1.Text);
                a.WriteLine(bunifuTextBox2.Text);
                a.WriteLine(bunifuTextBox3.Text);
                a.WriteLine(bunifuTextBox4.Text);
                a.WriteLine(bunifuTextBox5.Text);
                a.WriteLine(bunifuTextBox7.Text);
                a.WriteLine(bunifuTextBox13.Text);
                a.WriteLine(bunifuTextBox14.Text);
                a.WriteLine(bunifuTextBox6.Text);
                a.WriteLine(bunifuDatePicker1.Value);
                a.WriteLine(bunifuDatePicker2.Value);
            }
            Settings3.Default.ADIM = bunifuTextBox1.Text;
            Settings3.Default.SOYADIM = bunifuTextBox2.Text;
            Settings3.Default.TELEFONUM = bunifuTextBox3.Text;
            Settings3.Default.EPOSTA = bunifuTextBox4.Text;
            Settings3.Default.OKULUM = bunifuTextBox5.Text;
            Settings3.Default.ÖĞRETMEN = bunifuTextBox7.Text;
            Settings3.Default.BÖLÜM = bunifuTextBox13.Text;
            Settings3.Default.PROGRAM = bunifuTextBox14.Text;
            Settings3.Default.ADRES = bunifuTextBox6.Text;
            Settings3.Default.ŞİRKEADI = bunifuTextBox8.Text;
            Settings3.Default.SORUMLUKİŞİ = bunifuTextBox9.Text;
            Settings3.Default.ÜNVANI = bunifuTextBox10.Text;
            Settings3.Default.SORKİŞTEL = bunifuTextBox11.Text;
            Settings3.Default.ŞİRKETADRESİ = bunifuTextBox12.Text;
            if (bunifuCheckbox1.Checked == true)
            {
                Settings3.Default.HATIRLA = true;
                Settings3.Default.Save();
                MessageBox.Show("BİLGİLERİNİZ BAŞARIYLA KAYDEDİLMİŞTİR");
            }
            else
            {
                Settings3.Default.HATIRLA = false;
                Settings3.Default.Save();
                MessageBox.Show("BİLGİLERİNİZ BAŞARIYLA KAYDEDİLMİŞTİR");
            }
        }

        private void bunifuButton1_Click(object sender, EventArgs e)
        {
            
        }

        private void StajBilgilerim_Load(object sender, EventArgs e)
        {
            if (Settings3.Default.HATIRLA == true)
            {
                bunifuTextBox1.Text = Settings3.Default.ADIM;
                bunifuTextBox2.Text = Settings3.Default.SOYADIM;
                bunifuTextBox3.Text = Settings3.Default.TELEFONUM;
                bunifuTextBox4.Text = Settings3.Default.EPOSTA;
                bunifuTextBox5.Text = Settings3.Default.OKULUM;
                bunifuTextBox7.Text = Settings3.Default.ÖĞRETMEN;
                bunifuTextBox13.Text = Settings3.Default.BÖLÜM;
                bunifuTextBox14.Text = Settings3.Default.PROGRAM;
                bunifuTextBox6.Text = Settings3.Default.ADRES;
                bunifuDatePicker1.Text = Settings3.Default.BAŞTARİHİ;
                bunifuDatePicker2.Text = Settings3.Default.BAŞTARİHİ;
                bunifuTextBox8.Text = Settings3.Default.ŞİRKEADI;
                bunifuTextBox9.Text = Settings3.Default.SORUMLUKİŞİ;
                bunifuTextBox10.Text = Settings3.Default.ÜNVANI;
                bunifuTextBox11.Text = Settings3.Default.SORKİŞTEL;
                bunifuTextBox12.Text = Settings3.Default.ŞİRKETADRESİ;
            }
        }
    }
}

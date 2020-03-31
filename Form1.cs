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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (Settings1.Default.hatırla == true)
            {
                bunifuTextBox1.Text = Settings1.Default.id;
                bunifuTextBox2.Text = Settings1.Default.şifre;
                bunifuCheckbox1.Checked = Settings1.Default.hatırla;
            }
        }

        private void bunifuButton2_Click(object sender, EventArgs e)
        {

            if (bunifuTextBox3.Text==""& bunifuTextBox4.Text==""& bunifuTextBox5.Text==""& bunifuTextBox6.Text==""& bunifuTextBox3.Text==string.Empty||bunifuTextBox4.Text==string.Empty||bunifuTextBox5.Text==string.Empty||bunifuTextBox6.Text==string.Empty)
            {
                bunifuTextBox3.BorderColorIdle = Color.Yellow;
                bunifuTextBox4.BorderColorIdle = Color.Yellow;
                bunifuTextBox5.BorderColorIdle = Color.Yellow;
                bunifuTextBox6.BorderColorIdle = Color.Yellow;
                MessageBox.Show("LÜTFEN TÜM BİLGİLERİNİZİ EKSİKSİZ GİRİNİZ");
            }
            else
            {

                Directory.CreateDirectory(@"C:\ProgramData\Tenka\kayıt bilgisi");
                using (TextWriter a = new StreamWriter(@"C:\ProgramData\Tenka\kayıt bilgisi\kullanıcıadı.text")) ;
                using (TextWriter b = new StreamWriter(@"C:\ProgramData\Tenka\kayıt bilgisi\eposta.text")) ;
                using (TextWriter c = new StreamWriter(@"C:\ProgramData\Tenka\kayıt bilgisi\şifre.text")) ;
                using (TextWriter d = new StreamWriter(@"C:\ProgramData\Tenka\kayıt bilgisi\gsm.text")) ;

                using (StreamWriter ab = new StreamWriter(@"C:\ProgramData\Tenka\kayıt bilgisi\kullanıcıadı.text"))
                {
                    ab.WriteLine(bunifuTextBox3.Text);
                }
                using (StreamWriter ac = new StreamWriter(@"C:\ProgramData\Tenka\kayıt bilgisi\eposta.text"))
                {
                    ac.WriteLine(bunifuTextBox4.Text);
                }
                using (StreamWriter ad = new StreamWriter(@"C:\ProgramData\Tenka\kayıt bilgisi\şifre.text"))
                {
                    ad.WriteLine(bunifuTextBox5.Text);
                }
                using (StreamWriter ae = new StreamWriter(@"C:\ProgramData\Tenka\kayıt bilgisi\gsm.text"))
                {
                    ae.WriteLine(bunifuTextBox6.Text);
                }
                MessageBox.Show("KAYIT İŞLEMİNİZ BAŞARIYLA TAMAMLANDI");
                Settings1.Default.id = bunifuTextBox3.Text;
                Settings1.Default.şifre = bunifuTextBox5.Text;
                Settings1.Default.Save();
            }
            
        }

        private void bunifuButton1_Click(object sender, EventArgs e)
        {
            StreamReader a = new StreamReader(@"C:\ProgramData\Tenka\kayıt bilgisi\kullanıcıadı.text");
            string ad = a.ReadLine();
            StreamReader b = new StreamReader(@"C:\ProgramData\Tenka\kayıt bilgisi\şifre.text");
            string şi = b.ReadLine();

            if (bunifuTextBox1.Text == ad && bunifuTextBox2.Text == şi)
            {
                Öğrenci_Yönetim_Paneli öyp = new Öğrenci_Yönetim_Paneli();
                öyp.Show();
                Hide();
            }
            else
            {
                MessageBox.Show("HATALI ŞİFRE VEYA KULLANICI ADI GİRDİNİZ");
            }
            if (bunifuCheckbox1.Checked == true)
            {
                Settings1.Default.hatırla = true;
                Settings1.Default.Save();

                if (Settings1.Default.id == ad && Settings1.Default.şifre == şi)
                {
                    
                }
                else
             
             
            MessageBox.Show("KULLANICI ADI VEYA ŞİFRE HATALI");
            }
            else
            {
                Settings1.Default.hatırla = false;
                Settings1.Default.Save();
                if (Settings1.Default.id == ad && Settings1.Default.şifre == şi)
                {
                    
                }
                else
                {
                    MessageBox.Show("KULLANICI ADI VEYA ŞİFRE HATALI");
                }
            }
        }
        private void bunifuCheckbox1_OnChange(object sender, EventArgs e)
        {

        }

        private void bunifuGradientPanel1_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void bunifuTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

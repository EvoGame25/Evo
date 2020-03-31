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
    public partial class DersProgramı : Form
    {
        public DersProgramı()
        {
            InitializeComponent();
        }

        private void DersProgramı_Load(object sender, EventArgs e)
        {

        }

        private void bunifuButton1_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory(@"C:\ProgramData\Tenka\Ders Programı");
            using (TextWriter a = new StreamWriter(@"C:\ProgramData\Tenka\Ders Programı\pazartesi.text")) ;
            using (TextWriter b = new StreamWriter(@"C:\ProgramData\Tenka\Ders Programı\salı.text")) ;
            using (TextWriter c = new StreamWriter(@"C:\ProgramData\Tenka\Ders Programı\çarşamba.text"));
            using (TextWriter d = new StreamWriter(@"C:\ProgramData\Tenka\Ders Programı\perşembe.text")) ;
            using (TextWriter g = new StreamWriter(@"C:\ProgramData\Tenka\Ders Programı\cuma.text")) ;

            using(StreamWriter aa = new StreamWriter(@"C:\ProgramData\Tenka\Ders Programı\pazartesi.text"))
            {
                aa.WriteLine(bunifuTextBox5.Text);
                aa.WriteLine(maskedTextBox1.Text);
                aa.WriteLine(bunifuTextBox1.Text);
                aa.WriteLine(maskedTextBox2.Text);
                aa.WriteLine(bunifuTextBox2.Text);
                aa.WriteLine(maskedTextBox3.Text);
                aa.WriteLine(bunifuTextBox3.Text);
                aa.WriteLine(maskedTextBox4.Text);
                aa.WriteLine(bunifuTextBox4.Text);
                aa.WriteLine(maskedTextBox5.Text);
                aa.WriteLine(bunifuTextBox28.Text);
                aa.WriteLine(maskedTextBox6.Text);
            }
            using (StreamWriter bb = new StreamWriter(@"C:\ProgramData\Tenka\Ders Programı\salı.text"))
            {
                bb.WriteLine(bunifuTextBox6.Text);
                bb.WriteLine(maskedTextBox7.Text);
                bb.WriteLine(bunifuTextBox7.Text);
                bb.WriteLine(maskedTextBox8.Text);
                bb.WriteLine(bunifuTextBox8.Text);
                bb.WriteLine(maskedTextBox9.Text);
                bb.WriteLine(bunifuTextBox9.Text);
                bb.WriteLine(maskedTextBox10.Text);
                bb.WriteLine(bunifuTextBox10.Text);
                bb.WriteLine(maskedTextBox11.Text);
                bb.WriteLine(bunifuTextBox11.Text);
                bb.WriteLine(maskedTextBox12.Text);
            }
            using (StreamWriter cc = new StreamWriter(@"C:\ProgramData\Tenka\Ders Programı\çarşamba.text"))
            {
                cc.WriteLine(bunifuTextBox12.Text);
                cc.WriteLine(maskedTextBox13.Text);
                cc.WriteLine(bunifuTextBox13.Text);
                cc.WriteLine(maskedTextBox14.Text);
                cc.WriteLine(bunifuTextBox14.Text);
                cc.WriteLine(maskedTextBox15.Text);
                cc.WriteLine(bunifuTextBox15.Text);
                cc.WriteLine(maskedTextBox16.Text);
                cc.WriteLine(bunifuTextBox16.Text);
                cc.WriteLine(maskedTextBox17.Text);
                cc.WriteLine(bunifuTextBox17.Text);
                cc.WriteLine(maskedTextBox18.Text);
            }
            using (StreamWriter dd = new StreamWriter(@"C:\ProgramData\Tenka\Ders Programı\perşembe.text"))
            {
                dd.WriteLine(bunifuTextBox18.Text);
                dd.WriteLine(maskedTextBox19.Text);
                dd.WriteLine(bunifuTextBox19.Text);
                dd.WriteLine(maskedTextBox20.Text);
                dd.WriteLine(bunifuTextBox20.Text);
                dd.WriteLine(maskedTextBox21.Text);
                dd.WriteLine(bunifuTextBox21.Text);
                dd.WriteLine(maskedTextBox22.Text);
                dd.WriteLine(bunifuTextBox22.Text);
                dd.WriteLine(maskedTextBox23.Text);
                dd.WriteLine(bunifuTextBox29.Text);
                dd.WriteLine(maskedTextBox24.Text);
            }
            using (StreamWriter ee = new StreamWriter(@"C:\ProgramData\Tenka\Ders Programı\cuma.text"))
            {
                ee.WriteLine(bunifuTextBox23.Text);
                ee.WriteLine(maskedTextBox25.Text);
                ee.WriteLine(bunifuTextBox24.Text);
                ee.WriteLine(maskedTextBox26.Text);
                ee.WriteLine(bunifuTextBox25.Text);
                ee.WriteLine(maskedTextBox27.Text);
                ee.WriteLine(bunifuTextBox26.Text);
                ee.WriteLine(maskedTextBox28.Text);
                ee.WriteLine(bunifuTextBox27.Text);
                ee.WriteLine(maskedTextBox29.Text);
                ee.WriteLine(bunifuTextBox30.Text);
                ee.WriteLine(maskedTextBox30.Text);
            }
            MessageBox.Show("DERS PROGRAMINIZ BAŞARIYLA KAYDEDİLDİ");
        }

        private void bunifuButton2_Click(object sender, EventArgs e)
        {
            string[] pzrt = File.ReadAllLines(@"C:\ProgramData\Tenka\Ders Programı\pazartesi.text");
            if (pzrt.Length > 0)
            bunifuTextBox5.Text = pzrt[0];
            maskedTextBox1.Text = pzrt[1];
            bunifuTextBox1.Text = pzrt[2];
            maskedTextBox2.Text = pzrt[3];
            bunifuTextBox2.Text = pzrt[4];
            maskedTextBox3.Text = pzrt[5];
            bunifuTextBox3.Text = pzrt[6];
            maskedTextBox4.Text = pzrt[7];
            bunifuTextBox4.Text = pzrt[8];
            maskedTextBox5.Text = pzrt[9];
            bunifuTextBox28.Text = pzrt[10];
            maskedTextBox6.Text = pzrt[11];

            string[] salı = File.ReadAllLines(@"C:\ProgramData\Tenka\Ders Programı\salı.text");
            if (salı.Length > 0)
            bunifuTextBox6.Text = salı[0];
            maskedTextBox7.Text = salı[1];
            bunifuTextBox7.Text = salı[2];
            maskedTextBox8.Text = salı[3];
            bunifuTextBox8.Text = salı[4];
            maskedTextBox9.Text = salı[5];
            bunifuTextBox9.Text = salı[6];
            maskedTextBox10.Text = salı[7];
            bunifuTextBox10.Text = salı[8];
            maskedTextBox11.Text = salı[9];
            bunifuTextBox11.Text = salı[10];
            maskedTextBox12.Text = salı[11];

            string[] çrş = File.ReadAllLines(@"C:\ProgramData\Tenka\Ders Programı\çarşamba.text");
            if (çrş.Length > 0)
            bunifuTextBox12.Text = çrş[0];
            maskedTextBox13.Text = çrş[1];
            bunifuTextBox13.Text = çrş[2];
            maskedTextBox14.Text = çrş[3];
            bunifuTextBox14.Text = çrş[4];
            maskedTextBox15.Text = çrş[5];
            bunifuTextBox15.Text = çrş[6];
            maskedTextBox16.Text = çrş[7];
            bunifuTextBox16.Text = çrş[8];
            maskedTextBox17.Text = çrş[9];
            bunifuTextBox17.Text = çrş[10];
            maskedTextBox18.Text = çrş[11];

            string[] prş = File.ReadAllLines(@"C:\ProgramData\Tenka\Ders Programı\perşembe.text");
            if (prş.Length > 0)
            bunifuTextBox18.Text = prş[0];
            maskedTextBox19.Text = prş[1];
            bunifuTextBox19.Text = prş[2];
            maskedTextBox20.Text = prş[3];
            bunifuTextBox20.Text = prş[4];
            maskedTextBox21.Text = prş[5];
            bunifuTextBox21.Text = prş[6];
            maskedTextBox22.Text = prş[7];
            bunifuTextBox22.Text = prş[8];
            maskedTextBox23.Text = prş[9];
            bunifuTextBox29.Text = prş[10];
            maskedTextBox24.Text = prş[11];

            string[] cuma = File.ReadAllLines(@"C:\ProgramData\Tenka\Ders Programı\cuma.text");
            if (cuma.Length > 0)
            bunifuTextBox23.Text = cuma[0];
            maskedTextBox25.Text = cuma[1];
            bunifuTextBox24.Text = cuma[2];
            maskedTextBox26.Text = cuma[3];
            bunifuTextBox25.Text = cuma[4];
            maskedTextBox27.Text = cuma[5];
            bunifuTextBox26.Text = cuma[6];
            maskedTextBox28.Text = cuma[7];
            bunifuTextBox27.Text = cuma[8];
            maskedTextBox29.Text = cuma[9];
            bunifuTextBox30.Text = cuma[10];
            maskedTextBox30.Text = cuma[11];
        }
        
    }
}

using System;
using System.Windows.Forms;
using System.IO;



namespace TENKA_ÖĞRENCİ_PANELİ
{


    public partial class Öğrenci_Yönetim_Paneli : Form
    {

        public Öğrenci_Yönetim_Paneli()
        {
            InitializeComponent();
        }

            private void bunifuButton1_Click(object sender, EventArgs e)
        {
            Profilim prf = new Profilim();
            prf.Show();

        }

        private void bunifuButton6_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory(@"C:\ProgramData\kayıt bilgisi\photo");
            openFileDialog1.ShowDialog();
            bunifuPictureBox1.ImageLocation = openFileDialog1.FileName.ToString();
            using (StreamWriter streamWriter = new StreamWriter(@"C:\ProgramData\Tenka\kayıt bilgisi\photo"))
            {
                bunifuPictureBox1.Name = "ProfilPhoto";
                streamWriter.Write(bunifuPictureBox1.Image);

            }
        }

        private void bunifuButton3_Click(object sender, EventArgs e)
        {
            Notlarım ntl = new Notlarım();
            ntl.Show();

        }

        private void bunifuButton2_Click(object sender, EventArgs e)
        {
            DersProgramı drsp = new DersProgramı();
            drsp.Show();
        }

        private void bunifuButton5_Click(object sender, EventArgs e)
        {
            Hubis hbs = new Hubis();
            hbs.Show();
        }

        private void Öğrenci_Yönetim_Paneli_Load(object sender, EventArgs e)
        {

        }

        private void bunifuButton7_Click(object sender, EventArgs e)
        {
            StajBilgilerim stjbil = new StajBilgilerim();
            stjbil.Show();
        }

        private void bunifuButton8_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        protected virtual void _FormClosed(object sender, FormClosedEventArgs e)
        {

            DialogResult Cikis;
            Cikis = MessageBox.Show("Program Kapatılacak Emin siniz ?”, “Kapatma Uyarısı!" + MessageBoxButtons.YesNo);
            if (Cikis == DialogResult.Yes)
            {
                Application.Exit();
            }
            if (Cikis == DialogResult.No)
            {
                Application.Run();
            }

        }
        private void bunifuButton4_Click(object sender, EventArgs e)
        {
            Çalışmalarım çlşm = new Çalışmalarım();
            çlşm.Show();
        }

        private void bunifuButton9_Click(object sender, EventArgs e)
        {
            Officeİşlemlerim officeİşlemlerim = new Officeİşlemlerim();
            officeİşlemlerim.Show();
        }

        private void bunifuPictureBox1_Click(object sender, EventArgs e)
        {


        }
        
        private void bunifuTileButton1_Click(object sender, EventArgs e)
        {
            
        }
    }
}


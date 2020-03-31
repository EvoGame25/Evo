using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Text;
using Bunifu.Framework;
using System.IO;
using System.Text.RegularExpressions;
using Bunifu.Framework.UI;

namespace TENKA_ÖĞRENCİ_PANELİ
{



    public partial class Notlarım : Form
    {
        FontDialog fontDialog = new FontDialog();
        ColorDialog colorDialog = new ColorDialog();
        int tıklama = 0;

        public Notlarım()
        {
            InitializeComponent();

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void bunifuThinButton21_Click(object sender, EventArgs e)
        {


        }

        private void Notlarım_Load(object sender, EventArgs e)
        {
            if (Settings2.Default.hatırla == true)
            {
                bunifuMaterialTextbox1.Text = Settings2.Default.pazartesi;
                bunifuMaterialTextbox1.ForeColor = Settings2.Default.pazartesiRenk;
                bunifuMaterialTextbox2.Text = Settings2.Default.salı;
                bunifuMaterialTextbox2.ForeColor = Settings2.Default.salıRenk;
                bunifuMaterialTextbox3.Text = Settings2.Default.çarşamba;
                bunifuMaterialTextbox3.ForeColor = Settings2.Default.çarşambaRenk;
                bunifuMaterialTextbox4.Text = Settings2.Default.perşembe;
                bunifuMaterialTextbox4.ForeColor = Settings2.Default.perşembeRenk;
                bunifuMaterialTextbox5.Text = Settings2.Default.cuma;
                bunifuMaterialTextbox5.ForeColor = Settings2.Default.cumaRenk;
                bunifuMaterialTextbox6.Text = Settings2.Default.cumartesi;
                bunifuMaterialTextbox6.ForeColor = Settings2.Default.cumartesiRenk;
                bunifuMaterialTextbox7.Text = Settings2.Default.pazar;
                bunifuMaterialTextbox7.ForeColor = Settings2.Default.pazarRenk;
                bunifuMaterialTextbox8.Text = Settings2.Default.ekstra;
                bunifuMaterialTextbox8.ForeColor = Settings2.Default.ekstraRenk;
                bunifuMaterialTextbox9.Text = Settings2.Default.yeni1;
                bunifuMaterialTextbox9.ForeColor = Settings2.Default.yeni1Renk;
                bunifuMaterialTextbox10.Text = Settings2.Default.yeni2;
                bunifuMaterialTextbox10.ForeColor = Settings2.Default.yeni2Renk;
                bunifuMaterialTextbox11.Text = Settings2.Default.yeni3;
                bunifuMaterialTextbox11.ForeColor = Settings2.Default.yeni3Renk;
                bunifuMaterialTextbox12.Text = Settings2.Default.yeni4;
                bunifuMaterialTextbox12.ForeColor = Settings2.Default.yeni4Renk;
                bunifuCheckbox1.Checked = Settings2.Default.hatırla;
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void bunifuFlatButton1_Click(object sender, EventArgs e)
        {
            Settings2.Default.pazartesi = bunifuMaterialTextbox1.Text;
            Settings2.Default.pazartesiRenk = bunifuMaterialTextbox1.ForeColor;
            Settings2.Default.salı = bunifuMaterialTextbox2.Text;
            Settings2.Default.salıRenk = bunifuMaterialTextbox2.ForeColor;
            Settings2.Default.çarşamba = bunifuMaterialTextbox3.Text;
            Settings2.Default.çarşambaRenk = bunifuMaterialTextbox3.ForeColor;
            Settings2.Default.perşembe = bunifuMaterialTextbox4.Text;
            Settings2.Default.perşembeRenk = bunifuMaterialTextbox4.ForeColor;
            Settings2.Default.cuma = bunifuMaterialTextbox5.Text;
            Settings2.Default.cumaRenk = bunifuMaterialTextbox5.ForeColor;
            Settings2.Default.cumartesi = bunifuMaterialTextbox6.Text;
            Settings2.Default.cumartesiRenk = bunifuMaterialTextbox6.ForeColor;
            Settings2.Default.pazar = bunifuMaterialTextbox7.Text;
            Settings2.Default.pazarRenk = bunifuMaterialTextbox7.ForeColor;
            Settings2.Default.ekstra = bunifuMaterialTextbox8.Text;
            Settings2.Default.ekstraRenk = bunifuMaterialTextbox8.ForeColor;
            Settings2.Default.Save();
            if (bunifuCheckbox1.Checked == true)
            {
                Settings2.Default.hatırla = true;
                Settings2.Default.Save();
            }
            else
            {
                Settings2.Default.hatırla = false;
                Settings2.Default.Save();
            }
        }

        private void bunifuTextBox1_TextChanged(object sender, EventArgs e)
        {
            bunifuMaterialTextbox1.ForeColor = Color.White;
        }

        private void bunifuTextBox2_TextChanged(object sender, EventArgs e)
        {
            bunifuMaterialTextbox2.ForeColor = Color.White;
        }

        private void bunifuTextBox3_TextChanged(object sender, EventArgs e)
        {
            bunifuMaterialTextbox3.ForeColor = Color.White;
        }

        private void bunifuTextBox4_TextChanged(object sender, EventArgs e)
        {
            bunifuMaterialTextbox4.ForeColor = Color.White;
        }

        private void bunifuTextBox5_TextChanged(object sender, EventArgs e)
        {
            bunifuMaterialTextbox5.ForeColor = Color.White;
        }

        private void bunifuTextBox6_TextChanged(object sender, EventArgs e)
        {
            bunifuMaterialTextbox6.ForeColor = Color.White;
        }

        private void bunifuTextBox7_TextChanged(object sender, EventArgs e)
        {
            bunifuMaterialTextbox7.ForeColor = Color.White;
        }

        private void bunifuTextBox8_TextChanged(object sender, EventArgs e)
        {
            bunifuMaterialTextbox8.ForeColor = Color.White;
        }

        private void bunifuFlatButton10_Click(object sender, EventArgs e)
        {


            if (this.ActiveControl == bunifuMaterialTextbox1)
            {
                if (fontDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox1.Font = fontDialog.Font;
                }


            }
            if (this.ActiveControl == bunifuMaterialTextbox2)
            {
                if (fontDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox2.Font = fontDialog.Font;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox3)
            {
                if (fontDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox3.Font = fontDialog.Font;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox4)
            {
                if (fontDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox4.Font = fontDialog.Font;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox5)
            {
                if (fontDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox5.Font = fontDialog.Font;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox6)
            {
                if (fontDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox6.Font = fontDialog.Font;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox7)
            {
                if (fontDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox7.Font = fontDialog.Font;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox8)
            {
                if (fontDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox8.Font = fontDialog.Font;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox9)
            {
                if (fontDialog.ShowDialog()==DialogResult.OK)
                {
                    bunifuMaterialTextbox9.Font = fontDialog.Font;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox10)
            {
                if (fontDialog.ShowDialog()==DialogResult.OK)
                {
                    bunifuMaterialTextbox10.Font = fontDialog.Font;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox11)
            {
                if (fontDialog.ShowDialog()==DialogResult.OK)
                {
                    bunifuMaterialTextbox11.Font = fontDialog.Font;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox12)
            {
                if (fontDialog.ShowDialog()==DialogResult.OK)
                {
                    bunifuMaterialTextbox12.Font = fontDialog.Font;
                }
            }




        }

        private void bunifuFlatButton2_Click(object sender, EventArgs e)
        {


            if (this.ActiveControl == bunifuMaterialTextbox1)
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox1.ForeColor = colorDialog.Color;

                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox2)
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox2.ForeColor = colorDialog.Color;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox3)
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox3.ForeColor = colorDialog.Color;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox4)
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox4.ForeColor = colorDialog.Color;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox5)
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox5.ForeColor = colorDialog.Color;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox6)
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox6.ForeColor = colorDialog.Color;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox7)
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox7.ForeColor = colorDialog.Color;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox8)
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    bunifuMaterialTextbox8.ForeColor = colorDialog.Color;
                }
            }
            if (this.ActiveControl== bunifuMaterialTextbox9)
            {
                if (colorDialog.ShowDialog()==DialogResult.OK)
                {
                    bunifuMaterialTextbox9.ForeColor = colorDialog.Color;
                }
            }
            if (this.ActiveControl==bunifuMaterialTextbox10)
            {
                if (colorDialog.ShowDialog()==DialogResult.OK)
                {
                    bunifuMaterialTextbox10.ForeColor = colorDialog.Color;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox11)
            {
                if (colorDialog.ShowDialog()==DialogResult.OK)
                {
                    bunifuMaterialTextbox11.ForeColor = colorDialog.Color;
                }
            }
            if (this.ActiveControl == bunifuMaterialTextbox12)
            {
                if (colorDialog.ShowDialog()==DialogResult.OK)
                {
                    bunifuMaterialTextbox12.ForeColor = colorDialog.Color;
                }
            }
        }

        private void bunifuFlatButton11_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "WORD|*.doc , .docx|TXT|*.text";
            saveFileDialog1.Title = "BELGENİZİ KAYIT YERİNİ BELŞRLEYİN";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                StreamWriter streamWriter = new StreamWriter((System.IO.FileStream)saveFileDialog1.OpenFile());
                if (this.ActiveControl == bunifuMaterialTextbox1)
                {
                    streamWriter.WriteLine(bunifuMaterialTextbox1.Text, bunifuMaterialTextbox1.ForeColor);
                }
                if (this.ActiveControl == bunifuMaterialTextbox2)
                {
                    streamWriter.WriteLine(bunifuMaterialTextbox2.Text);
                }
                if (this.ActiveControl == bunifuMaterialTextbox3)
                {
                    streamWriter.WriteLine(bunifuMaterialTextbox3.Text);
                }
                if (this.ActiveControl == bunifuMaterialTextbox4)
                {
                    streamWriter.WriteLine(bunifuMaterialTextbox4.Text);
                }
                if (this.ActiveControl == bunifuMaterialTextbox5)
                {
                    streamWriter.WriteLine(bunifuMaterialTextbox5.Text);
                }
                if (this.ActiveControl == bunifuMaterialTextbox6)
                {
                    streamWriter.WriteLine(bunifuMaterialTextbox6.Text);
                }
                if (this.ActiveControl == bunifuMaterialTextbox7)
                {
                    streamWriter.WriteLine(bunifuMaterialTextbox7.Text);
                }
                if (this.ActiveControl == bunifuMaterialTextbox8)
                {
                    streamWriter.WriteLine(bunifuMaterialTextbox8.Text);
                }
                if (this.ActiveControl == bunifuMaterialTextbox9)
                {
                    streamWriter.WriteLine(bunifuMaterialTextbox9.Text);
                }
                if (this.ActiveControl == bunifuMaterialTextbox10)
                {
                    streamWriter.WriteLine(bunifuMaterialTextbox10.Text);
                }
                if (this.ActiveControl == bunifuMaterialTextbox11)
                {
                    streamWriter.WriteLine(bunifuMaterialTextbox11.Text);
                }
                if (this.ActiveControl == bunifuMaterialTextbox12)
                {
                    streamWriter.WriteLine(bunifuMaterialTextbox12.Text);
                }
                streamWriter.Close();
            }
        }

        private void bunifuFlatButton3_Click(object sender, EventArgs e)
        {
            if (this.ActiveControl == bunifuMaterialTextbox1)
            {
                bunifuMaterialTextbox1.Text = "";

            }
            if (this.ActiveControl == bunifuMaterialTextbox2)
            {
                bunifuMaterialTextbox2.Text = "";
            }
            if (this.ActiveControl == bunifuMaterialTextbox3)
            {
                bunifuMaterialTextbox3.Text = "";
            }
            if (this.ActiveControl == bunifuMaterialTextbox4)
            {
                bunifuMaterialTextbox4.Text = "";
            }
            if (this.ActiveControl == bunifuMaterialTextbox5)
            {
                bunifuMaterialTextbox5.Text = "";
            }
            if (this.ActiveControl == bunifuMaterialTextbox6)
            {
                bunifuMaterialTextbox6.Text = "";
            }
            if (this.ActiveControl == bunifuMaterialTextbox7)
            {
                bunifuMaterialTextbox7.Text = "";
            }
            if (this.ActiveControl == bunifuMaterialTextbox8)
            {
                bunifuMaterialTextbox8.Text = "";
            }
            if (this.ActiveControl == bunifuMaterialTextbox9)
            {
                bunifuMaterialTextbox9.Text = "";
            }
            if (this.ActiveControl == bunifuMaterialTextbox10)
            {
                bunifuMaterialTextbox10.Text = "";
            }
            if (this.ActiveControl==bunifuMaterialTextbox11)
            {
                bunifuMaterialTextbox11.Text = "";
            }
            if (this.ActiveControl==bunifuMaterialTextbox12)
            {
                bunifuMaterialTextbox12.Text = "";
            }
        }

        private void bunifuFlatButton7_Click(object sender, EventArgs e)
        {
            if (this.ActiveControl == bunifuMaterialTextbox1)
            {
                bunifuMaterialTextbox1.TextAlign = HorizontalAlignment.Left;
            }
            if (this.ActiveControl == bunifuMaterialTextbox2)
            {
                bunifuMaterialTextbox2.TextAlign = HorizontalAlignment.Left;
            }
            if (this.ActiveControl == bunifuMaterialTextbox3)
            {
                bunifuMaterialTextbox3.TextAlign = HorizontalAlignment.Left;
            }
            if (this.ActiveControl == bunifuMaterialTextbox4)
            {
                bunifuMaterialTextbox4.TextAlign = HorizontalAlignment.Left;
            }
            if (this.ActiveControl == bunifuMaterialTextbox5)
            {
                bunifuMaterialTextbox5.TextAlign = HorizontalAlignment.Left;
            }
            if (this.ActiveControl == bunifuMaterialTextbox6)
            {
                bunifuMaterialTextbox6.TextAlign = HorizontalAlignment.Left;
            }
            if (this.ActiveControl == bunifuMaterialTextbox7)
            {
                bunifuMaterialTextbox7.TextAlign = HorizontalAlignment.Left;
            }
            if (this.ActiveControl == bunifuMaterialTextbox8)
            {
                bunifuMaterialTextbox8.TextAlign = HorizontalAlignment.Left;
            }
            if (this.ActiveControl == bunifuMaterialTextbox9)
            {
                bunifuMaterialTextbox9.TextAlign = HorizontalAlignment.Left;
            }
            if (this.ActiveControl == bunifuMaterialTextbox10)
            {
                bunifuMaterialTextbox10.TextAlign = HorizontalAlignment.Left;
            }
            if (this.ActiveControl == bunifuMaterialTextbox11)
            {
                bunifuMaterialTextbox11.TextAlign = HorizontalAlignment.Left;
            }
            if (this.ActiveControl == bunifuMaterialTextbox12)
            {
                bunifuMaterialTextbox12.TextAlign = HorizontalAlignment.Left;
            }
        }

        private void bunifuFlatButton8_Click(object sender, EventArgs e)
        {
            if (this.ActiveControl == bunifuMaterialTextbox1)
            {
                bunifuMaterialTextbox1.TextAlign = HorizontalAlignment.Center;
            }
            if (this.ActiveControl == bunifuMaterialTextbox2)
            {
                bunifuMaterialTextbox2.TextAlign = HorizontalAlignment.Center;
            }
            if (this.ActiveControl == bunifuMaterialTextbox3)
            {
                bunifuMaterialTextbox3.TextAlign = HorizontalAlignment.Center;
            }
            if (this.ActiveControl == bunifuMaterialTextbox4)
            {
                bunifuMaterialTextbox4.TextAlign = HorizontalAlignment.Center;
            }
            if (this.ActiveControl == bunifuMaterialTextbox5)
            {
                bunifuMaterialTextbox5.TextAlign = HorizontalAlignment.Center;
            }
            if (this.ActiveControl == bunifuMaterialTextbox6)
            {
                bunifuMaterialTextbox6.TextAlign = HorizontalAlignment.Center;
            }
            if (this.ActiveControl == bunifuMaterialTextbox7)
            {
                bunifuMaterialTextbox7.TextAlign = HorizontalAlignment.Center;
            }
            if (this.ActiveControl == bunifuMaterialTextbox8)
            {
                bunifuMaterialTextbox8.TextAlign = HorizontalAlignment.Center;
            }
            if (this.ActiveControl == bunifuMaterialTextbox9)
            {
                bunifuMaterialTextbox9.TextAlign = HorizontalAlignment.Center;
            }
            if (this.ActiveControl == bunifuMaterialTextbox10)
            {
                bunifuMaterialTextbox10.TextAlign = HorizontalAlignment.Center;
            }
            if (this.ActiveControl == bunifuMaterialTextbox11)
            {
                bunifuMaterialTextbox11.TextAlign = HorizontalAlignment.Center;
            }
            if (this.ActiveControl == bunifuMaterialTextbox12)
            {
                bunifuMaterialTextbox12.TextAlign = HorizontalAlignment.Center;
            }
        }

        private void bunifuFlatButton9_Click(object sender, EventArgs e)
        {
            if (this.ActiveControl == bunifuMaterialTextbox1)
            {
                bunifuMaterialTextbox1.TextAlign = HorizontalAlignment.Right;
            }
            if (this.ActiveControl == bunifuMaterialTextbox2)
            {
                bunifuMaterialTextbox2.TextAlign = HorizontalAlignment.Right;
            }
            if (this.ActiveControl == bunifuMaterialTextbox3)
            {
                bunifuMaterialTextbox3.TextAlign = HorizontalAlignment.Right;
            }
            if (this.ActiveControl == bunifuMaterialTextbox4)
            {
                bunifuMaterialTextbox4.TextAlign = HorizontalAlignment.Right;
            }
            if (this.ActiveControl == bunifuMaterialTextbox5)
            {
                bunifuMaterialTextbox5.TextAlign = HorizontalAlignment.Right;
            }
            if (this.ActiveControl == bunifuMaterialTextbox6)
            {
                bunifuMaterialTextbox6.TextAlign = HorizontalAlignment.Right;
            }
            if (this.ActiveControl == bunifuMaterialTextbox7)
            {
                bunifuMaterialTextbox7.TextAlign = HorizontalAlignment.Right;
            }
            if (this.ActiveControl == bunifuMaterialTextbox8)
            {
                bunifuMaterialTextbox8.TextAlign = HorizontalAlignment.Right;
            }
            if (this.ActiveControl == bunifuMaterialTextbox9)
            {
                bunifuMaterialTextbox9.TextAlign = HorizontalAlignment.Right;
            }
            if (this.ActiveControl == bunifuMaterialTextbox10)
            {
                bunifuMaterialTextbox10.TextAlign = HorizontalAlignment.Right;
            }
            if (this.ActiveControl == bunifuMaterialTextbox11)
            {
                bunifuMaterialTextbox11.TextAlign = HorizontalAlignment.Right;
            }
            if (this.ActiveControl == bunifuMaterialTextbox12)
            {
                bunifuMaterialTextbox12.TextAlign = HorizontalAlignment.Right;
            }
        }

        private void bunifuMaterialTextbox1_OnValueChanged(object sender, EventArgs e)
        {

        }

        private void bunifuFlatButton4_Click(object sender, EventArgs e)
        {
            if (tıklama <= 4)
            {
                tıklama++;
            }

            BunifuFlatButton bunifuFlatButton = new BunifuFlatButton();
            bunifuFlatButton.Click += new System.EventHandler(this.bunifuFlatButton4_Click);
            Controls.Add(bunifuFlatButton);

            if (tıklama == 1)
            {
                bunifuMaterialTextbox9.Visible = true;
            }

            if (tıklama == 2)
            {
                bunifuMaterialTextbox10.Visible = true;
            }
            if (tıklama == 3)
            {
                bunifuMaterialTextbox11.Visible = true;
                
            }
            if (tıklama == 4)
            {
                bunifuMaterialTextbox12.Visible = true;
            }
        }
    }
}

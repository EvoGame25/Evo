using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TENKA_ÖĞRENCİ_PANELİ
{
    
    public partial class TENKA : Form
    {
        List<Color> colors = new List<Color>();
        public TENKA()
        {
            colors.Add(Color.FromArgb(0, 158, 71));
            colors.Add(Color.FromArgb(112, 191, 83));
            colors.Add(Color.FromArgb(216, 155, 40));
            colors.Add(Color.FromArgb(217, 102, 41));
            colors.Add(Color.FromArgb(217, 102, 41));
            colors.Add(Color.FromArgb(235, 83, 104));
            colors.Add(Color.FromArgb(223, 128, 255));
            colors.Add(Color.FromArgb(112, 48, 160));
            colors.Add(Color.FromArgb(107, 122, 187));
            colors.Add(Color.FromArgb(95, 136, 176));
            colors.Add(Color.FromArgb(70, 175, 227));
            colors.Add(Color.FromArgb(0, 158, 71));
            
            
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
        int curcolor = 0;
        int loop = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            if (curcolor<colors.Count -1)
            {
                this.BackColor = Bunifu.Framework.UI.BunifuColorTransition.getColorScale(loop, colors[curcolor], colors[curcolor + 1]);
                if (loop<100)
                {
                    loop++;
                    
                }
                else
                {
                    loop = 0;
                    curcolor++;
                }
                timer1.Enabled = true;
            }
            else
            {
                Form1 frm1 = new Form1();
                frm1.Show();
                Hide();

            }
        }
        
        private void timer2_Tick(object sender, EventArgs e)
        {
            timer2.Start();
            if (curcolor<colors.Count -1)
            {
                this.BackColor = Bunifu.Framework.UI.BunifuColorTransition.getColorScale(loop, colors[curcolor], colors[curcolor + 1]);
                if (loop<100)
                {
                    loop++;
                    bunifuCircleProgressbar1.ProgressBackColor = colors[curcolor + 1];
                    bunifuCircleProgressbar1.Value += 10;
                    
                }
                else
                {
                    loop = 0;
                    curcolor++;
                }
                timer1.Enabled = true;
            }
            
            

            
        }
    }
}

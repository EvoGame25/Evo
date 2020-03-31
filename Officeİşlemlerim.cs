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
    public partial class Officeİşlemlerim : Form
    {
        public Officeİşlemlerim()
        {
            InitializeComponent();
        }
        private void bunifuTileButton1_Click(object sender, EventArgs e)
        {
            Wordİşlermleri wordİşlermleri = new Wordİşlermleri();
            wordİşlermleri.Show();
        }
    }
}

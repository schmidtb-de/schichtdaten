using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Schichtdaten
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            tbdickemin.Text = Properties.Settings.Default.DickeMin;
            tbdickemax.Text = Properties.Settings.Default.DickeMax;
            tbgewichtmin.Text = Properties.Settings.Default.GewichtMin;
            tbgewichtmax.Text = Properties.Settings.Default.GewichtMax;
            tbbreitemin.Text = Properties.Settings.Default.BreiteMin;
            tbbreitemax.Text = Properties.Settings.Default.BreiteMax;
            tbausbauchungmin.Text = Properties.Settings.Default.AusbauchungMin;
            tbausbauchungmax.Text = Properties.Settings.Default.AusbauchungMax;
        }

        private void toleranzenspeichern_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.DickeMin = tbdickemin.Text;
            Properties.Settings.Default.DickeMax = tbdickemax.Text;
            Properties.Settings.Default.GewichtMin = tbgewichtmin.Text;
            Properties.Settings.Default.GewichtMax = tbgewichtmax.Text;
            Properties.Settings.Default.BreiteMin = tbbreitemin.Text;
            Properties.Settings.Default.BreiteMax = tbbreitemax.Text;
            Properties.Settings.Default.AusbauchungMin = tbausbauchungmin.Text;
            Properties.Settings.Default.AusbauchungMax = tbausbauchungmax.Text;
            Properties.Settings.Default.Save();
            this.Close();
        }
    }
}

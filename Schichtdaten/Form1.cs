using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using System.Collections.Specialized;
using System.Collections;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;


namespace Schichtdaten
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            mesautobutton.FlatAppearance.BorderColor = System.Drawing.Color.WhiteSmoke;

            Application.EnableVisualStyles();

            string subfolder = System.IO.Path.Combine(Application.StartupPath, "Stundeneingabe");
            System.IO.Directory.CreateDirectory(subfolder);
            string subfoldersave = System.IO.Path.Combine(subfolder, "save");
            System.IO.Directory.CreateDirectory(subfoldersave);
            string sinoxpfadsave = subfoldersave + "\\SINOxPfad.txt";
            if (System.IO.File.Exists(sinoxpfadsave))
            {
                txtSinoxPfad.Text = System.IO.File.ReadAllText(sinoxpfadsave);
                txtSinoxPfad.Enabled = false;
            }


            System.IO.DirectoryInfo ParentDirectory = new System.IO.DirectoryInfo(subfolder);

            foreach (System.IO.FileInfo f in ParentDirectory.GetFiles())
            {
                string woextension = Path.GetFileNameWithoutExtension(f.Name);
                string User = (woextension.Substring(0, 8));
                string Group = (woextension.Substring(8));

                if (Group == "Hauck")
                {
                    listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[0]));
                }

                if (Group == "Körner")
                {
                    listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[1]));
                }
                if (Group == "Diel")
                {
                    listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[2]));
                }
                if (Group == "Werner")
                {
                    listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[3]));
                }
                if (Group == "Beier")
                {
                    listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[4]));
                }
                if (Group == "Moser")
                {
                    listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[5]));
                }
                if (Group == "Springer")
                {
                    listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[6]));
                }


            }

            
                

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Datum.Text = DateTime.Now.ToString("dd.MM.yyyy");
            txtEingabeAPlatten.Text = Properties.Settings.Default.PlattenAnzahl;
            ErechnerZielAnzahl.Text = Properties.Settings.Default.ZielAnzahl;
            MrechnerSchnittCharge.Text = Properties.Settings.Default.DurchschnittCharge;
            ErechnerSchnittSchicht.Text = Properties.Settings.Default.DurchschnittSchicht;
            ErechnerAktuell.Text = Properties.Settings.Default.AktuellAnzahl;
            this.TopMost = true;
            button2.Enabled = false;            
            button1.Enabled = false; 
            comboBox1.Items.AddRange(new string[] { "3 Mann", "2 Mann" });
            comboBox1.SelectedItem = "3 Mann";
            lblLaufzeit.Text = "0 - 0 min" ;
            coBoAnlage.Items.AddRange(new string[] { "Beschichtung 2", "Beschichtung 3"});
            coBoAnlage.SelectedItem = "Beschichtung 3";
            if (coBoAnlage.SelectedItem == "Beschichtung 3")
            {
                coBoSchicht.Items.Clear();
                coBoSchicht.Items.AddRange(new string[] { "Hauck", "Körner", "Diel", "Springer" });
                coBoSchicht.SelectedItem = "Hauck";
            }
            else
            {
                coBoSchicht.Items.Clear();
                coBoSchicht.Items.AddRange(new string[] { "Werner", "Beier", "Moser", "Springer" });
                coBoSchicht.SelectedItem = "Werner";
            }
            string path = Application.StartupPath + "\\EElement.txt";
            if (System.IO.File.Exists(path))
                txtEingabeEElement.Text = System.IO.File.ReadAllText(path);
            txtEingabeAusschuss.Select();


        }

        private void lblGesAusschuss_Anzeige_Click(object sender, EventArgs e)
        {

        }

        private void txtErstesElement_TextChanged(object sender, EventArgs e)
        {
            //int EElementgleich1;
            //bool success1 = Int32.TryParse(txtEingabeEElement.Text, out EElementgleich1);
            if (txtEingabeEElement.Text == "1")
            {
                lblEEkleiner1.Text = "<- Plattenanzahl richtig?";
            }
            else
            {
                lblEEkleiner1.Text = "";
            }
        }

        private void txtEingabeAusschuss_TextChanged(object sender, EventArgs e)
        {
            Datum.Text = DateTime.Now.ToString("dd.MM.yyyy");
            
        }

        private void txtAusgabeElemente_TextChanged(object sender, EventArgs e)
        {
            
            
            
        } 

        private void button1_Click(object sender, EventArgs e)
        {
            Datum.Text = DateTime.Now.ToString("dd.MM.yyyy");
            for (int i = 1; i <= 6; i++)
            {
                if (Properties.Settings.Default.Pruefzahl == "1")
                {
                    Properties.Settings.Default.Schicht1 = txtAusgabeElemente.Text;
                    Properties.Settings.Default.Pruefzahl = "2";
                    Properties.Settings.Default.Save();
                    break;
                }
                if (Properties.Settings.Default.Pruefzahl == "2")
                {
                    Properties.Settings.Default.Schicht2 = txtAusgabeElemente.Text;
                    Properties.Settings.Default.Pruefzahl = "3";
                    Properties.Settings.Default.Save();
                    break;
                }
                if (Properties.Settings.Default.Pruefzahl == "3")
                {
                    Properties.Settings.Default.Schicht3 = txtAusgabeElemente.Text;
                    Properties.Settings.Default.Pruefzahl = "4";
                    Properties.Settings.Default.Save();
                    break;
                }
                if (Properties.Settings.Default.Pruefzahl == "4")
                {
                    Properties.Settings.Default.Schicht4 = txtAusgabeElemente.Text;
                    Properties.Settings.Default.Pruefzahl = "5";
                    Properties.Settings.Default.Save();
                    break;
                }
                if (Properties.Settings.Default.Pruefzahl == "5")
                {
                    Properties.Settings.Default.Schicht5 = txtAusgabeElemente.Text;
                    Properties.Settings.Default.Pruefzahl = "6";
                    Properties.Settings.Default.Save();
                    break;
                }
                if (Properties.Settings.Default.Pruefzahl == "6")
                {
                    Properties.Settings.Default.Schicht6 = txtAusgabeElemente.Text;
                    Properties.Settings.Default.Pruefzahl = "1";
                    Properties.Settings.Default.Save();
                    break;
                }
            }

            //lblElemente_Vorher.Text = txtAusgabeElemente.Text;
            //lblPlatten_Vorher.Text = txtAusgabePlatten.Text;
            //lblAusschuss_Vorher.Text = txtAusgabeAusschuss.Text;
            txtEingabeEElement.Text = txtEingabeLElement.Text;
            txtEingabeLElement.Text = "";
            txtEingabeAusschuss.Text = "";
            txtAusgabeElemente.Text = "";
            txtAusgabePlatten.Text = "";
            txtAusgabeAusschuss.Text = "";
            lblEEkleiner1.Text = "";
            string path = Application.StartupPath + "\\EElement.txt";
            System.IO.File.WriteAllText(path, txtEingabeEElement.Text);
            txtEingabeAusschuss.Select();
            ErechnerAktuell.Text = Convert.ToString((Convert.ToInt32(txtEingabeEElement.Text))-1);
            int S1 = Convert.ToInt32(Properties.Settings.Default.Schicht1);
            int S2 = Convert.ToInt32(Properties.Settings.Default.Schicht2);
            int S3 = Convert.ToInt32(Properties.Settings.Default.Schicht3);
            int S4 = Convert.ToInt32(Properties.Settings.Default.Schicht4);
            int S5 = Convert.ToInt32(Properties.Settings.Default.Schicht5);
            int S6 = Convert.ToInt32(Properties.Settings.Default.Schicht6);
            int Ergebnis = ((S1 + S2 + S3 + S4 + S5 + S6) / 6);
            ErechnerSchnittSchicht.Text = Convert.ToString(Ergebnis);

                   }


        private void txtEingabeAPlatten_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.PlattenAnzahl = txtEingabeAPlatten.Text;
            Properties.Settings.Default.Save();

        }

        private void txtEingabeLElement_TextChanged(object sender, EventArgs e)
        {
            
           
                bool success1;
                bool success2;
                bool success3;
                int EElement;
                int LElement;
                int Plattenanzahl;
                success1 = Int32.TryParse(txtEingabeEElement.Text, out EElement);
                success2 = Int32.TryParse(txtEingabeLElement.Text, out LElement);
                success3 = Int32.TryParse(txtEingabeAPlatten.Text, out Plattenanzahl);
                /*int EElement = Convert.ToInt32(txtEingabeEElement.Text);
                int LElement = Convert.ToInt32(txtEingabeLElement.Text);
                int Plattenanzahl = Convert.ToInt32(txtEingabeAPlatten.Text);*/
                if (LElement > EElement)
                {
                    button1.Enabled = true;
                    int intElemente = LElement - EElement;
                    int intPlattenanzahl = intElemente * Plattenanzahl;
                    String sPlattenanzahl = Convert.ToString(intPlattenanzahl);
                    String Elemente = Convert.ToString(intElemente);
                    txtAusgabeElemente.Text = Elemente;
                    txtAusgabePlatten.Text = sPlattenanzahl;

                    double int3Mann1 = intPlattenanzahl / 7.2 / 2;
                    decimal d1 = (decimal)int3Mann1;                    // Runden
                    decimal d3Mann1 = Math.Round(d1);                   // Runden
                    double int3Mann2 = intPlattenanzahl / 7.0 / 2;
                    decimal d2 = (decimal)int3Mann2;
                    decimal d3Mann2 = Math.Round(d2);
                    double int2Mann1 = intPlattenanzahl / 6.2 / 2;
                    decimal d3 = (decimal)int2Mann1;
                    decimal d2Mann1 = Math.Round(d3);
                    double int2Mann2 = intPlattenanzahl / 6.0 / 2;
                    decimal d4 = (decimal)int2Mann2;
                    decimal d2Mann2 = Math.Round(d4);

                    String s3Mann1 = Convert.ToString(d3Mann1);
                    String s3Mann2 = Convert.ToString(d3Mann2);
                    String s2Mann1 = Convert.ToString(d2Mann1);
                    String s2Mann2 = Convert.ToString(d2Mann2);
                    if (comboBox1.SelectedIndex == 0)
                        lblLaufzeit.Text = s3Mann1 + " - " + s3Mann2 + " min";
                    else
                        lblLaufzeit.Text = s2Mann1 + " - " + s2Mann2 + " min";
                }
                else
                {
                    button1.Enabled = false;
                    txtAusgabeElemente.Text = "";
                    txtAusgabePlatten.Text = "";
                    lblLaufzeit.Text = "0 - 0 min";
                }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
        }

        private void txtAusgabePlatten_TextChanged(object sender, EventArgs e)
        {
            
       
        }

        private void lblPlatten_Vorher_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool success1;
            bool success2;
            bool success3;
            int EElement;
            int LElement;
            int Plattenanzahl;
            success1 = Int32.TryParse(txtEingabeEElement.Text, out EElement);
            success2 = Int32.TryParse(txtEingabeLElement.Text, out LElement);
            success3 = Int32.TryParse(txtEingabeAPlatten.Text, out Plattenanzahl);
            /*int EElement = Convert.ToInt32(txtEingabeEElement.Text);
            int LElement = Convert.ToInt32(txtEingabeLElement.Text);
            int Plattenanzahl = Convert.ToInt32(txtEingabeAPlatten.Text);*/
            int intElemente = LElement - EElement;
            int intPlattenanzahl = intElemente * Plattenanzahl;
            String sPlattenanzahl = Convert.ToString(intPlattenanzahl);
            String Elemente = Convert.ToString(intElemente);

            double int3Mann1 = intPlattenanzahl / 7.2 / 2;
            decimal d1 = (decimal)int3Mann1;                    // Runden
            decimal d3Mann1 = Math.Round(d1);                   // Runden
            double int3Mann2 = intPlattenanzahl / 7.0 / 2;
            decimal d2 = (decimal)int3Mann2;
            decimal d3Mann2 = Math.Round(d2);
            double int2Mann1 = intPlattenanzahl / 6.2 / 2;
            decimal d3 = (decimal)int2Mann1;
            decimal d2Mann1 = Math.Round(d3);
            double int2Mann2 = intPlattenanzahl / 6.0 / 2;
            decimal d4 = (decimal)int2Mann2;
            decimal d2Mann2 = Math.Round(d4);

            String s3Mann1 = Convert.ToString(d3Mann1);
            String s3Mann2 = Convert.ToString(d3Mann2);
            String s2Mann1 = Convert.ToString(d2Mann1);
            String s2Mann2 = Convert.ToString(d2Mann2);
            if (txtEingabeLElement.Text != "")
            {
                if (comboBox1.SelectedIndex == 0)
                    lblLaufzeit.Text = s3Mann1 + " - " + s3Mann2 + " min";
                else
                    lblLaufzeit.Text = s2Mann1 + " - " + s2Mann2 + " min";
            }
            txtEingabeAusschuss.Select();
        }

        private void txtEingabeLElement_MouseClick(object sender, MouseEventArgs e)
        {
            txtEingabeLElement.SelectAll();
        }

        private void txtEingabeLElement_Enter(object sender, EventArgs e)
        {
            txtEingabeLElement.SelectAll();
        }

        private void txtEingabeAusschuss_Enter(object sender, EventArgs e)
        {
            txtEingabeAusschuss.SelectAll();
        }

        private void txtEingabeAPlatten_Enter(object sender, EventArgs e)
        {
            txtEingabeAPlatten.SelectAll();
        }

        private void lblLaufzeit_Click(object sender, EventArgs e)
        {

        }

        private void lblLaufzeit_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

            if (checkBox1.Checked)
                this.TopMost = true;
            else
                this.TopMost = false;
            txtEingabeAusschuss.Select();
               
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (txtBenutzername.Text == "" & txtPasswort.Text == "")
            {
                button2.Enabled = false;
            }


            string SinoxPfad = "shell.run \"\"\"" + txtSinoxPfad.Text + "\"\"\"\n";
            string Benutzer = "shell.SendKeys \"" + txtBenutzername.Text + "\"\n";
            string Passwort = "shell.SendKeys \"" + txtPasswort.Text + "\"\n";
            string Pfad = AppDomain.CurrentDomain.BaseDirectory + txtBenutzername + ".vbs";
            string[] lines = { "Public Function WaitForWindow(WindowTitle)\n",
"Set WshShell = WScript.CreateObject(\"WScript.Shell\")\n", 
"success=0\n",
"I = 0\n",
"Do\n",
"I = I + 1\n",
"WScript.Sleep 300\n",
"success=WshShell.AppActivate(WindowTitle)\n",
"If I = 50 Then\n",
"x=MsgBox (\"Fenster \"+Chr(34)& WindowTitle &Chr(34)+\" konnte nicht gefunden werden.\",48,\"Warnmeldung\")\n",
"WaitForWindow = False\n",
"asyncConnection.Disconnect(2)\n",
"WScript.quit\n",
"Exit Do\n",
"End If\n",
"Loop Until success\n",
"If success Then\n",
"WaitForWindow = True\n",
"End If\n",
"End Function\n",
"'Deklaration\n",
"set shell = CreateObject(\"WScript.Shell\")\n",
"'Ausführen des Programmes\n",
//SinoxPfad,
"WScript.Sleep 500\n",
"shell.AppActivate \"SINOx-Info\"\n",
"'Tastendruck simulieren\n",
"WaitForWindow(\"Anmelden\")\n",
"WScript.Sleep 100\n",
Benutzer,
"WScript.Sleep 100\n",
"shell.SendKeys \"{TAB}\"\n",
Passwort,
"WScript.Sleep 100\n",
"shell.SendKeys \"{ENTER}\"\n",
"WaitForWindow(\"SQL Server Login\")\n",
"WScript.Sleep 100\n",
"shell.SendKeys \"{TAB}\"\n",
"WScript.Sleep 100\n",
Benutzer,
"WScript.Sleep 100\n",
"shell.SendKeys \"{TAB}\"\n",
"WScript.Sleep 100\n",
Passwort,
"WScript.Sleep 100\n",
"shell.SendKeys \"{TAB}\"\n", 
"WScript.Sleep 100\n",
"shell.SendKeys \"{ENTER}\"\n",
"WaitForWindow(\"Anmelden\")\n",
"WScript.Sleep 100\n",
Benutzer,
"WScript.Sleep 100\n",
"shell.SendKeys \"{TAB}\"\n",
"WScript.Sleep 100\n",
Passwort,
"WScript.Sleep 100\n",
"shell.SendKeys \"{ENTER}\"\n",
"WaitForWindow(\"SQL Server Login\")\n",
"WScript.Sleep 100\n",
"shell.SendKeys \"{TAB}\"\n",
"WScript.Sleep 100\n",
Benutzer,
"WScript.Sleep 100\n",
"shell.SendKeys \"{TAB}\"\n",
"WScript.Sleep 100\n",
Passwort,
"WScript.Sleep 100\n",
"shell.SendKeys \"{TAB}\"\n", 
"WScript.Sleep 100\n",
"shell.SendKeys \"{ENTER}\"\n",
"WScript.Sleep 5000\n",
Passwort,
"WScript.Sleep 100\n",
"shell.SendKeys \"{ENTER}\"\n",
"WaitForWindow(\"SQL Server Login\")\n",
"WScript.Sleep 100\n",
"shell.SendKeys \"{TAB}\"",
"WScript.Sleep 100",
Benutzer,
"WScript.Sleep 100",
"shell.SendKeys \"{TAB}\"",
"WScript.Sleep 100",
Passwort,
"WScript.Sleep 100",
"shell.SendKeys \"{TAB}\"", 
"WScript.Sleep 100",
"shell.SendKeys \"{ENTER}\"",


"WScript.Sleep 1500",
"shell.SendKeys \"{TAB}\"",
"shell.SendKeys \"{RIGHT}\"",
"shell.SendKeys \"{RIGHT}\"",
"shell.SendKeys \" \"",
                            
                             };


            if (txtBenutzername.Text.Length == 8)
            {
                string subfolder = System.IO.Path.Combine(Application.StartupPath, "Stundeneingabe");
                System.IO.Directory.CreateDirectory(subfolder);
                string pathsave = subfolder + "\\" + txtBenutzername.Text + coBoSchicht.SelectedItem.ToString() + ".vbs";
                System.IO.File.WriteAllLines(pathsave, lines);
                string subfoldersave = System.IO.Path.Combine(subfolder, "save");
                System.IO.Directory.CreateDirectory(subfoldersave);
                string sinoxpfadsave = subfoldersave + "\\SINOxPfad.txt";
                System.IO.File.WriteAllText(sinoxpfadsave, txtSinoxPfad.Text);
                txtSinoxPfad.Enabled = false;

                txtBenutzername.Text = "";
                txtPasswort.Text = "";
                listView1.Items.Clear();
                System.IO.DirectoryInfo ParentDirectory = new System.IO.DirectoryInfo(subfolder);

                foreach (System.IO.FileInfo f in ParentDirectory.GetFiles())
                {
                    string woextension = Path.GetFileNameWithoutExtension(f.Name);
                    string User = (woextension.Substring(0, 8));
                    string Group = (woextension.Substring(8));

                    if (Group == "Hauck")
                    {
                        listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[0]));
                    }

                    if (Group == "Körner")
                    {
                        listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[1]));
                    }
                    if (Group == "Diel")
                    {
                        listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[2]));
                    }
                    if (Group == "Werner")
                    {
                        listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[3]));
                    }
                    if (Group == "Beier")
                    {
                        listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[4]));
                    }
                    if (Group == "Moser")
                    {
                        listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[5]));
                    }
                    if (Group == "Springer")
                    {
                        listView1.Items.Add(new ListViewItem(new string[] { User }, listView1.Groups[6]));
                    }
                }
            }
            else
            {
                MessageBox.Show("Der Benutzername muss genau 8 Stellen haben!", "Fehler");
            }

        }

        private void lsbBenutzer_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            string Ausgewählt = listView1.SelectedItems.ToString();
            string subfolder = System.IO.Path.Combine(Application.StartupPath, "Stundeneingabe");
            // Create the subfolder
            System.IO.Directory.CreateDirectory(subfolder);
            System.IO.DirectoryInfo ParentDirectory = new System.IO.DirectoryInfo(subfolder);

            foreach (System.IO.FileInfo f in ParentDirectory.GetFiles())
            {
                string woextension = Path.GetFileNameWithoutExtension(f.Name);
                string User = (woextension.Substring(0, 8));
                string Group = (woextension.Substring(8));

                if (Ausgewählt == User)
                {
                    string pathsave = subfolder + "\\" + woextension + ".vbs";
                    System.Diagnostics.Process process = new System.Diagnostics.Process();
                    process.EnableRaisingEvents = false;
                    process.StartInfo.FileName = pathsave;
                    process.Start();
                }
                
            }
            
            

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void cmdMinus_Click(object sender, EventArgs e)
        {

            bool success2;
            int AusschussGesamt;
            success2 = Int32.TryParse(txtAusgabeAusschuss.Text, out AusschussGesamt);
            if (AusschussGesamt != 0)
            {
                int AusschussGesamtNeu = AusschussGesamt - 1;
                String AuschussGesamt = Convert.ToString(AusschussGesamtNeu);
                txtAusgabeAusschuss.Text = AuschussGesamt;
            }
            else
                txtAusgabeAusschuss.Text = "0";
            txtEingabeAusschuss.Select();
            
        }

        private void cmdPlus_Click(object sender, EventArgs e)
        {

            bool success1;
            bool success2;
            int NeuerAusschuss;
            int AusschussGesamt;
            success1 = Int32.TryParse(txtEingabeAusschuss.Text, out NeuerAusschuss);
            success2 = Int32.TryParse(txtAusgabeAusschuss.Text, out AusschussGesamt);           
            int AusschussGesamtNeu = AusschussGesamt + NeuerAusschuss;
            String AuschussGesamt = Convert.ToString(AusschussGesamtNeu);
            txtEingabeAusschuss.Text = "";            
            txtAusgabeAusschuss.Text = AuschussGesamt;
            txtEingabeAusschuss.Select();

        }

        private void coBoAnlage_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (coBoAnlage.SelectedItem == "Beschichtung 3")
            {
                coBoSchicht.Items.Clear();
                coBoSchicht.Items.AddRange(new string[] { "Hauck", "Körner", "Diel", "Springer" });
                coBoSchicht.SelectedItem = "Hauck";
            }
            else
            {
                coBoSchicht.Items.Clear();
                coBoSchicht.Items.AddRange(new string[] { "Werner", "Beier", "Moser", "Springer" });
                coBoSchicht.SelectedItem = "Werner";
            }
        }

        private void txtBenutzername_TextChanged(object sender, EventArgs e)
        {
            if (txtBenutzername.Text != "" & txtPasswort.Text != "")
            {
                button2.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
            }
        }

        private void txtPasswort_TextChanged(object sender, EventArgs e)
        {
            if (txtBenutzername.Text != "" & txtPasswort.Text != "")
            {
                button2.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
            }
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
          
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            string Ausgewählt = listView1.SelectedItems[0].Text + listView1.SelectedItems[0].Group.Name;
            //Create a new subfolder under the current active folder
            string subfolder = System.IO.Path.Combine(Application.StartupPath, "Stundeneingabe");
            // Create the subfolder
            System.IO.Directory.CreateDirectory(subfolder);
            //string pathsave = Application.StartupPath +"\\" + txtBenutzername.Text + ".vbs";
            string pathsave = subfolder + "\\" + Ausgewählt + ".vbs";
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = @txtSinoxPfad.Text;
            proc.Start();
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.EnableRaisingEvents = false;
            process.StartInfo.FileName = pathsave;
            process.Start();
            
           
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                txtEingabeAusschuss.Select();
            }
            if (tabControl1.SelectedIndex == 1)
            {
                txtBenutzername.Select();
            }
        }

        private void mesautobutton_Click(object sender, EventArgs e)
        {
        }

        private void mesautobutton_MouseDown(object sender, MouseEventArgs e)
        {
            int Prüf1;
            int Prüf2;
            if (txtEingabeEElement.Text == "")
            Prüf1 = 0;
            else
            Prüf1 = Convert.ToInt32(txtEingabeEElement.Text);
            if (txtEingabeLElement.Text == "")
                Prüf2 = 0;
            else
            Prüf2 = Convert.ToInt32(txtEingabeLElement.Text);

                if (e.Button == MouseButtons.Right)
                {
                    txtEingabeAusschuss.Select();
                    if (Prüf1 - Prüf2 == 5)
                    new Form2().Show(this);

                }
                if (e.Button == MouseButtons.Left)
                {
                    if (Prüf1 - Prüf2 == 5)
                    {
                        int Dicke1 = Convert.ToInt32(Properties.Settings.Default.DickeMin);
                        int Dicke2 = Convert.ToInt32(Properties.Settings.Default.DickeMax);
                        Random rnd = new Random();
                        int ZufallszahlDicke1 = rnd.Next(Dicke1, Dicke2);
                        int ZufallszahlDicke2 = rnd.Next(Dicke1, Dicke2);
                        int ZufallszahlDicke3 = rnd.Next(Dicke1, Dicke2);
                        int Gewicht1 = Convert.ToInt32(Properties.Settings.Default.GewichtMin);
                        int Gewicht2 = Convert.ToInt32(Properties.Settings.Default.GewichtMax);
                        Random rnd2 = new Random();
                        int ZufallszahlGewicht = rnd2.Next(Gewicht1, Gewicht2);
                        int Breite1 = Convert.ToInt32(Properties.Settings.Default.BreiteMin);
                        int Breite2 = Convert.ToInt32(Properties.Settings.Default.BreiteMax);
                        Random rnd3 = new Random();
                        int ZufallszahlBreite = rnd3.Next(Breite1, Breite2);
                        int Ausbauchung1 = Convert.ToInt32(Properties.Settings.Default.AusbauchungMin);
                        int Ausbauchung2 = Convert.ToInt32(Properties.Settings.Default.AusbauchungMax);
                        Random rnd4 = new Random();
                        int ZufallszahlAusbauchung = rnd4.Next(Ausbauchung1, Ausbauchung2);
                        String meskeys1 = "," + ZufallszahlDicke1.ToString() + " {TAB} " + "," + ZufallszahlDicke2.ToString() + " {TAB} " + "," + ZufallszahlDicke3.ToString() + " {TAB} " + ZufallszahlGewicht.ToString() + " {Tab} " + ZufallszahlBreite.ToString() + " {TAB} " + ZufallszahlAusbauchung.ToString() + " {TAB} " + " {TAB} " + " {TAB}" + "{ENTER}";

                        Interaction.AppActivate(@"Johnson Matthey MES - powered by Xavo - [Beschichtungsanlage-3]");
                        System.Threading.Thread.Sleep(100);
                        string[] keys = meskeys1.Split(' ');
                        foreach (string key in keys)
                        {
                            SendKeys.Send(key);
                        }
                    }
                    txtEingabeAusschuss.Select();
                }
            
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void ErechnerAktuell_TextChanged(object sender, EventArgs e)
        {
            int Aktuell;
            int andere;
            int Ziel;
            int Rest;
            if (ErechnerAktuell.Text == "")
                Aktuell = 0;
            else
                Aktuell = Convert.ToInt32(ErechnerAktuell.Text);
            if (ErechnerAktuellaA.Text == "")
                andere = 0;
            else
                andere = Convert.ToInt32(ErechnerAktuellaA.Text);
            if (ErechnerZielAnzahl.Text == "")
                Ziel = 0;
            else
                Ziel = Convert.ToInt32(ErechnerZielAnzahl.Text);
            //Rest = Convert.ToInt32(ErechnerRestAnzahl.Text);
            if (Ziel - (Aktuell + andere) < 0)
                ErechnerRestAnzahl.Text = "";
            else
            ErechnerRestAnzahl.Text = Convert.ToString(Ziel - (Aktuell + andere));
            Properties.Settings.Default.AktuellAnzahl = ErechnerAktuell.Text;
            Properties.Settings.Default.Save();

        }

        private void ErechnerZielAnzahl_TextChanged(object sender, EventArgs e)
        {
            int Aktuell;
            int andere;
            int Ziel;
            int Rest;
            if (ErechnerAktuell.Text == "")
                Aktuell = 0;
            else
                Aktuell = Convert.ToInt32(ErechnerAktuell.Text);
            if (ErechnerAktuellaA.Text == "")
                andere = 0;
            else
                andere = Convert.ToInt32(ErechnerAktuellaA.Text);
            if (ErechnerZielAnzahl.Text == "")
                Ziel = 0;
            else
                Ziel = Convert.ToInt32(ErechnerZielAnzahl.Text);
            //Rest = Convert.ToInt32(ErechnerRestAnzahl.Text);
            if (Ziel - (Aktuell + andere) < 0)
                ErechnerRestAnzahl.Text = "";
            else
                ErechnerRestAnzahl.Text = Convert.ToString(Ziel - (Aktuell + andere));
            Properties.Settings.Default.ZielAnzahl = ErechnerZielAnzahl.Text;
            Properties.Settings.Default.Save();
       
        }

        private void ErechnerAktuellaA_TextChanged(object sender, EventArgs e)
        {
            int Aktuell;
            int andere;
            int Ziel;
            int Rest;
            if (ErechnerAktuell.Text == "")
                Aktuell = 0;
            else
                Aktuell = Convert.ToInt32(ErechnerAktuell.Text);
            if (ErechnerAktuellaA.Text == "")
                andere = 0;
            else
                andere = Convert.ToInt32(ErechnerAktuellaA.Text);
            if (ErechnerZielAnzahl.Text == "")
                Ziel = 0;
            else
                Ziel = Convert.ToInt32(ErechnerZielAnzahl.Text);
            //Rest = Convert.ToInt32(ErechnerRestAnzahl.Text);
            if (Ziel - (Aktuell + andere) < 0)
                ErechnerRestAnzahl.Text = "";
            else
                ErechnerRestAnzahl.Text = Convert.ToString(Ziel - (Aktuell + andere));
        }

        private void MrechnerSchnittCharge_TextChanged(object sender, EventArgs e)
        {
            int Rest;
            int SchnittCh;
            if (ErechnerRestAnzahl.Text == "")
                Rest = 0;
            else
                Rest = Convert.ToInt32(ErechnerRestAnzahl.Text);
            if (ErechnerRestAnzahl.Text == "")
                MrechnerFehlendeCh.Text = "";
            else
            {
                SchnittCh = Convert.ToInt32(MrechnerSchnittCharge.Text);
                MrechnerFehlendeCh.Text = Convert.ToString(Rest / SchnittCh);
            }
            Properties.Settings.Default.DurchschnittCharge = MrechnerSchnittCharge.Text;
            Properties.Settings.Default.Save();
        }

        private void ErechnerRestAnzahl_TextChanged(object sender, EventArgs e)
        {
            int Rest;
            int SchnittCh;
            int SchnittEl;
            if (ErechnerRestAnzahl.Text == "")
            {
                Rest = 0;
            }
            else
            {
                Rest = Convert.ToInt32(ErechnerRestAnzahl.Text);
                if (ErechnerSchnittSchicht.Text == "")
                    SchnittEl = 0;
                else
                {
                    SchnittEl = Convert.ToInt32(ErechnerSchnittSchicht.Text);
                    VerblSchichten.Text = Convert.ToString(Rest / SchnittEl);
                }
            }
            if (ErechnerRestAnzahl.Text == "")
            {
                MrechnerFehlendeCh.Text = "";
                VerblSchichten.Text = "";
                VerblTage.Text = "";
                VerblDatum.Text = "";
            }
            else
            {
                if (MrechnerSchnittCharge.Text == "")
                    MrechnerFehlendeCh.Text = "0";
                else
                {
                    SchnittCh = Convert.ToInt32(MrechnerSchnittCharge.Text);
                    MrechnerFehlendeCh.Text = Convert.ToString(Rest / SchnittCh);
                }
            }
        }

        private void VerblSchichten_TextChanged(object sender, EventArgs e)
        {
            int Schichten;
            if (VerblSchichten.Text == "")
                VerblTage.Text = "";
            else
            {
                Schichten = Convert.ToInt32(VerblSchichten.Text);
                VerblTage.Text = Convert.ToString(Schichten / 3);
            }
        }

        private void VerblTage_TextChanged(object sender, EventArgs e)
        {
            int Tage;
            if (VerblTage.Text == "")
                VerblDatum.Text = "";
            else
            {
                Tage = Convert.ToInt32(VerblTage.Text);
                for (int i = 1; i <= 4; i++)
                {
                    if (Tage >= 7 && Tage < 14)
                    {
                        Tage = Tage + 2;
                        break;
                    }
                    if (Tage >= 14 && Tage < 21)
                    {
                        Tage = Tage + 4;
                        break;
                    }
                    if (Tage >= 21 && Tage < 28)
                    {
                        Tage = Tage + 6;
                        break;
                    }
                    if (Tage >= 28)
                    {
                        Tage = Tage + 8;
                        break;
                    }
                }
                VerblDatum.Text = DateTime.Now.AddDays(Tage).ToString("dd.MM.yyyy");
            }

        }

        private void ErechnerSchnittSchicht_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.DurchschnittSchicht = ErechnerSchnittSchicht.Text;
            Properties.Settings.Default.Save();
        }
    

       
    }
}

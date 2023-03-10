using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Main = Meter.MyApplicationContext;

namespace Meter.Forms
{
    public partial class OpenArchive : Form
    {
        string thisYear, thisMonth;
        string archPath;
        private Dictionary<string, string> archMap;

        public OpenArchive(string year, string month)
        {
            thisYear = year;
            thisMonth = month;
            archMap = new Dictionary<string, string>();
            InitializeComponent();
        }

        private void btn_Click(object sender, EventArgs e)
        {
            string sourceFolder = Main.dir + @"\current";
            Control c = sender as Control;
            if (c.Text == thisMonth)
            {
                return;
            }
            NewMenuBase.month = c.Text;
            NewMenuBase.year = textBox1.Text;

            Main.instance.wb.Save();
            Main.instance.Arhivate(thisYear, thisMonth);
            Main.instance.wb.Close();
            Directory.Delete(sourceFolder, true);
            Directory.CreateDirectory(sourceFolder);
            System.IO.Compression.ZipFile.ExtractToDirectory(archMap[c.Text], sourceFolder);
            Main.instance.Restart();
        }

        private void OpenArchive_Load(object sender, EventArgs e)
        {
            this.textBox1.Text = DateTime.Now.ToString("yyyy");
            Check();
        }

        private void Check()
        {
            archMap.Clear();
            archPath = Main.dir + @"\arch";
            if (!Directory.Exists(archPath))
            {
                DisableAll();
                return;
                //Directory.CreateDirectory(archPath);
            }
            archPath = archPath + @"\" + this.textBox1.Text;
            if (!Directory.Exists(archPath))
            {
                DisableAll();
                return;
                //Directory.CreateDirectory(archPath);
            }
            foreach (Control item in flowLayoutPanel1.Controls)
            {
                string arhiveName = archPath + @"\" + item.Text + @".zip";
                if (File.Exists(arhiveName))
                {
                    archMap.Add(item.Text, arhiveName);
                    item.Enabled = true;
                    if (item.Text == thisMonth)
                    {
                        item.BackColor = Color.Gold;
                    }
                    else
                    {
                        item.BackColor = Color.GreenYellow;
                    }
                }
                else
                {
                    item.Enabled = false;
                    item.BackColor = SystemColors.ButtonFace;
                }
            }
        }

        private void DisableAll()
        {
            foreach (Control item in flowLayoutPanel1.Controls)
            {
                item.Enabled = false;
                item.BackColor = SystemColors.ButtonFace;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if (!Char.IsDigit(c) && c != 8)
            {
                e.Handled = true;
            }
        }

        private void OpenArchive_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                Main.Exit();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Check();
        }
    }
}

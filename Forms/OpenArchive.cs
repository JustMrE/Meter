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
    public partial class OpenArchive : MyFormBase
    {
        string thisYear, thisMonth;
        string selectedYear, selectedMonth;
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
            GlobalMethods.ToLog(this, sender);
            Control c = sender as Control;
            if (c.BackColor == Color.GreenYellow)
            {
                return;
            }
            else
            {
                selectedMonth = c.Text;
            }
            NewMenuBase.month = selectedMonth;
            NewMenuBase.year = selectedYear;

            GlobalMethods.ToLog("Открытие архива за " + selectedMonth + " " + selectedYear);

            //Main.instance.RunOnUiThread(Main.instance.Restart, thisYear, thisMonth, archMap[c.Text]);
            //Main.instance.RunOnUiThread(Main.instance.OpenMonth, thisYear, thisMonth, selectedMonth, archMap[c.Text]);
            Main.instance.OpenMonth(thisYear, thisMonth, selectedYear, selectedMonth, archMap[c.Text]);
            Close();
        }

        private void OpenArchive_Load(object sender, EventArgs e)
        {
            this.textBox1.Text = thisYear;
        }

        private void Check()
        {
            archMap.Clear();
            archPath = MeterSettings.DBDir + @"\arch";
            if (!Directory.Exists(archPath))
            {
                DisableAll();
                return;
                //Directory.CreateDirectory(archPath);
            }
            archPath = archPath + @"\" + selectedYear;
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
                    if (selectedYear == thisYear && item.Text == thisMonth)
                    {
                        item.BackColor = Color.GreenYellow;
                        item.ForeColor = SystemColors.ControlDark;
                    }
                    else
                    {
                        item.BackColor = Color.Gold;
                        item.ForeColor = SystemColors.ControlText;
                    }
                }
                else
                {
                    item.Enabled = false;
                    item.BackColor = SystemColors.ButtonFace;
                    item.ForeColor = SystemColors.ControlText;
                }
            }
        }

        private void DisableAll()
        {
            foreach (Control item in flowLayoutPanel1.Controls)
            {
                item.Enabled = false;
                item.BackColor = SystemColors.ButtonFace;
                item.ForeColor = SystemColors.ControlText;
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

        }

        private void OpenArchive_Shown(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            selectedYear = ((TextBox)sender).Text;
            GlobalMethods.ToLog(this, sender, selectedYear);
            Check();
        }

        public void FormClose()
        {
            Action action = () => {System.Windows.Forms.Application.ExitThread(); };
            if (InvokeRequired)
            {
                Invoke(action);
            }
            else
            {
                action();
            }
        }
    }
}

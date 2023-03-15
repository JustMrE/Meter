using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Meter.Forms
{
    public partial class YearSelect : Form
    {
        public string? year;
        public YearSelect(string thisYear)
        {
            InitializeComponent();
            this.textBox1.Text = thisYear;
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if (!Char.IsDigit(c) && c != 8)
            {
                e.Handled = true;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            year = null;
            Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Введите год!");
                return;
            }
            year = textBox1.Text;
            GlobalMethods.ToLog("Изменен год на " + year);
            Close();
        }
    }
}

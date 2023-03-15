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
    public partial class MonthSelect : Form
    {
        public string? selectedMonth = null;
        public MonthSelect()
        {
            InitializeComponent();
        }

        private void btn_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            Control c = sender as Control;

            selectedMonth = c.Text;

            GlobalMethods.ToLog("Изменен месяц на " + selectedMonth);

            Close();
        }

        private void MonthSelect_Shown(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this);
        }
    }
}

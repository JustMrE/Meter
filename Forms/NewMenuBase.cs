using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;

namespace Meter.Forms
{
    public partial class NewMenuBase : Form
    {
        

        public NewMenuBase()
        {
            InitializeComponent();
        }

        protected virtual void NewMenuBase_Load(object sender, EventArgs e)
        {
            listBox1.Items.AddRange(Main.instance.references.references.Keys.OrderBy(m => m).ToArray());
            formHwnd = this.Handle;
            SetParent(formHwnd, Main.instance.xlAppHwnd);

            listBox1.DoubleClick += new EventHandler(listBox1_DoubleClick);
            tbSearch.TextChanged += new EventHandler(searchTextBox_Changed);
            checkBox1.CheckedChanged += new EventHandler(searchTextBox_Changed);
        }
        protected virtual void NewMenuBase_Shown(object sender, EventArgs e)
        {
            SetWindowLong(this.Handle, GWL_EXSTYLE, WS_EX_TOOLWINDOW);
            Main.instance.menu = this;
            GlobalMethods.CalculateFormsPositions();
        }
        protected virtual void NewMenuBase_Closing(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }
        protected virtual void NewMenuBase_Activated(object sender, EventArgs e)
        {
            GlobalMethods.CalculateFormsPositions();
        }
        private void MenuBase_FormClosed(object sender, FormClosedEventArgs e)
        {
            ResetContextMenu();
            Marshal.ReleaseComObject(cb);
            if (_activeRange != null) Marshal.ReleaseComObject(_activeRange);
        }
        protected virtual void btnAdmin_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button42_Click(object sender, EventArgs e)
        {

        }
        protected virtual void RepairMenu_Click(object sender, EventArgs e)
        {
            GlobalMethods.CalculateFormsPositions();
        }
        protected virtual void listBox1_DoubleClick(object sender, EventArgs e)
        {
            string codeName = ((Excel.Worksheet)Main.instance.xlApp.ActiveSheet).CodeName;
            if (codeName == "PS")
            {
                Main.instance.references[(string)listBox1.SelectedItem].PS.Range.Select();
                SetForegroundWindow(Main.instance.xlAppHwnd);

            }
            else if (codeName == "DB")
            {
                Main.instance.references[(string)listBox1.SelectedItem].DB.Range.Select();
                SetForegroundWindow(Main.instance.xlAppHwnd);
            }
        }
        protected virtual void Button2_Click(object sender, EventArgs e)
        {

        }
        protected virtual void TextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
    }
}

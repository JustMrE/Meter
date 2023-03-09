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
using System.IO.Compression;
using System.Globalization;

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
        protected virtual void TextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        protected virtual void Button2_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button3_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button4_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button5_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button6_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button7_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button8_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button9_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button10_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button11_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button12_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button13_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button14_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button15_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button16_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button17_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button18_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button19_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button20_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button21_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button22_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button23_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button24_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button25_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button26_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button27_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button28_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button29_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button30_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button31_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button32_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button33_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button34_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button35_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button36_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button37_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button38_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button39_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button40_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button41_Click(object sender, EventArgs e)
        {

        }
        protected virtual void Button42_Click(object sender, EventArgs e)
        {

        }
    
        protected virtual void btnToArhive_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Внимание!\nЕсли архив уже существует, он будет перезаписан! Вы хотите продолжить?", caption: "Предупреждение!", MessageBoxButtons.OKCancel ,icon: MessageBoxIcon.Exclamation) == DialogResult.Cancel)
            {
                return;
            }

            string sourceFolder = Main.dir + @"\current";
            string tempDirectory = Path.Combine(Path.GetTempPath(), DateTime.Today.ToString("MMMM", new CultureInfo("ru-RU")));
            Directory.CreateDirectory(tempDirectory);
            foreach (string dirPath in Directory.GetDirectories(sourceFolder, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourceFolder, tempDirectory));
            }
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*", SearchOption.AllDirectories))
            {
                File.Copy(filePath, filePath.Replace(sourceFolder, tempDirectory), true);
            }
            string archPath = Main.dir + @"\arch";
            if (!Directory.Exists(archPath))
            {
                Directory.CreateDirectory(archPath);
            }
            archPath = archPath + @"\" + this.lblYear.Text;
            if (!Directory.Exists(archPath))
            {
                Directory.CreateDirectory(archPath);
            }
            string arhiveName = archPath + @"\" + this.lblMonth.Text + @".zip";
            if (File.Exists(arhiveName))
            {
                File.Delete(arhiveName);
            }
            ZipFile.CreateFromDirectory(tempDirectory, arhiveName);
            Directory.Delete(tempDirectory, true);
            MessageBox.Show("Готово!");
        }
    }
}

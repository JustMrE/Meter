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
using static System.Net.Mime.MediaTypeNames;

namespace Meter.Forms
{
    public partial class NewMenuBase : Form
    {
        public static string? month = null, year = null;

        public NewMenuBase()
        {
            InitializeComponent();
            this.lblMonth.Text = Main.instance.wsCh.Range["B5"].Value.ToString();
            this.lblYear.Text = Main.instance.wsCh.Range["D5"].Value.ToString();
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
            GlobalMethods.ReleseObject(cb);
            GlobalMethods.ReleseObject(_activeRange);
        }
        protected virtual void NewMenuBase_Activated(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(month)) this.lblMonth.Text = month;
            if (!string.IsNullOrEmpty(year)) this.lblYear.Text = year;
            GlobalMethods.CalculateFormsPositions();
        }
        private void MenuBase_FormClosed(object sender, FormClosedEventArgs e)
        {
            ResetContextMenu();
            GlobalMethods.ReleseObject(cb);
            GlobalMethods.ReleseObject(_activeRange);
        }


        protected virtual void btnAdmin_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void RepairMenu_Click(object sender, EventArgs e)
        {
            ToLog(sender);
            GlobalMethods.CalculateFormsPositions();
        }
        protected virtual void listBox1_DoubleClick(object sender, EventArgs e)
        {
            ToLog(sender, listBox1.SelectedItem);
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
        protected virtual void searchTextBox_Changed(object sender, EventArgs e)
        {
            ToLog(sender, this.tbSearch.Text);
            RegexSearch();
        }
        protected virtual void TextBox1_TextChanged(object sender, EventArgs e)
        {
            ToLog(sender, this.textBox1.Text);
            if (textBox1.Text == "stop")
            {
                ResetContextMenu();
                Main.instance.xlApp.EnableEvents = false;
            }
            else
            {
                Main.instance.xlApp.EnableEvents = true;
            }
            if (textBox1.Text == "dontsave")
            {
                Main.dontsave = true;
            }
            else
            {
                Main.dontsave = false;
            }
        }

        protected virtual void ToLog(object sender)
        {
            GlobalMethods.ToLog("Нажато " + ((Control)sender).Name);
        }
        protected virtual void ToLog(object sender, string txt)
        {
            GlobalMethods.ToLog("Изменен текст " + ((Control)sender).Name + " на '" + txt + "'");
        }
        protected virtual void ToLog(object sender, object selectedItem)
        {
            GlobalMethods.ToLog("В списке " + ((Control)sender).Name + " выбран " + selectedItem.ToString());
        }

        protected virtual void lblMonth_TextChanged(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender, lblMonth.Text);
            Main.instance.StopAll();
            Main.instance.wsCh.Range["B5"].Value = lblMonth.Text;
            Main.instance.ResumeAll();
        }

        protected virtual void lblYear_TextChanged(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender, lblYear.Text);
            Main.instance.StopAll();
            Main.instance.wsCh.Range["D5"].Value = lblYear.Text;
            Main.instance.ResumeAll();
        }

        protected virtual void Button2_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button3_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button4_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button5_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button6_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button7_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button8_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button9_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button10_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button11_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button12_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button13_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button14_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button15_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button16_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button17_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button18_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button19_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button20_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button21_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button22_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button23_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button24_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button25_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button26_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button27_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button28_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button29_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button30_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button31_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button32_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button33_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button34_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button35_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button36_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button37_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button38_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button39_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button40_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button41_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
        protected virtual void Button42_Click(object sender, EventArgs e)
        {
            ToLog(sender);
        }
    
        protected virtual void btnToArhive_Click(object sender, EventArgs e)
        {
            ToLog(sender);
            if (MessageBox.Show("Внимание!\nЕсли архив уже существует, он будет перезаписан! Вы хотите продолжить?", caption: "Предупреждение!", MessageBoxButtons.OKCancel ,icon: MessageBoxIcon.Exclamation) == DialogResult.Cancel)
            {
                return;
            }

            Main.instance.ArhivateNew(this.lblYear.Text, this.lblMonth.Text);

            MessageBox.Show("Готово!");
        }

        protected virtual void btnFromArhive_Click(object sender, EventArgs e)
        {
            ToLog(sender);
            // Thread t = new Thread(() =>
            // {
            //     OpenArchive form = new OpenArchive(this.lblYear.Text, this.lblMonth.Text);
            //     form.FormClosed += (s, args) =>
            //     {
            //         System.Windows.Forms.Application.ExitThread();
            //     };
            //     form.Show();
            //     System.Windows.Forms.Application.Run();
            // });
            // t.SetApartmentState(ApartmentState.STA);
            // t.Start();
            OpenArchive form = new OpenArchive(this.lblYear.Text, this.lblMonth.Text);
            form.ShowDialog();
        }

        protected virtual void lblMonth_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ToLog(sender);
        }

        protected virtual void lblYear_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ToLog(sender);
        }
    }
}

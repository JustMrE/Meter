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
using Excel = Microsoft.Office.Interop.Excel;

namespace Meter.Forms
{
    public partial class TempForm : Form
    {
        public static TempForm instance;
        public NewMenuBase menuForm;
        public List<Form> menues;

        public TempForm()
        {
            InitializeComponent();
            if (instance == null) instance = this;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ShowNewMenu();
        }

        private void ShowNewMenu()
        {
            menues = new List<Form>();
            menues.Add(new NewMenu());
            menues.Add(new NewMenuAdmin());
            menuForm = menues[0] as NewMenuBase;
            menuForm.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (Form form in menues)
            {
                form.Close();
            }
        }

        public void CalculateFormsPositions()
        {
            Main.instance.StopAll();
            if (Main.instance.xlApp.WindowState != Excel.XlWindowState.xlMinimized && !Main.instance.closed)
            {
                Main.GetWindowRect(Main.instance.xlAppHwnd, ref Main.instance.xlAppRect);
                Main.instance.xlAppRect.Top = Main.instance.xlAppRect.Top < 0 ? 0 : Main.instance.xlAppRect.Top;
                Main.instance.xlAppRect.Left = Main.instance.xlAppRect.Left < 0 ? 0 : Main.instance.xlAppRect.Left;
                string range = "B7";
                Excel.Range r = ((Excel.Range)Main.instance.xlApp.Selection);
                string oldRange = r.Address;
                int scrollValueR = Main.instance.xlApp.ActiveWindow.ScrollRow;
                int scrollValueC = Main.instance.xlApp.ActiveWindow.ScrollColumn;
                if (Main.instance.xlApp.ActiveWindow.FreezePanes == true)
                {
                    Main.instance.xlApp.ActiveWindow.FreezePanes = true;
                    Main.instance.xlApp.ActiveWindow.FreezePanes = false;
                }
                int cellPosX = (int)Main.instance.xlApp.ActiveWindow.PointsToScreenPixelsX(0);
                int cellPosY = (int)Main.instance.xlApp.ActiveWindow.PointsToScreenPixelsY(0);

                int left = cellPosX - Main.instance.xlAppRect.Left;
                int top = cellPosY - Main.instance.xlAppRect.Top;
                int width = Main.instance.xlAppRect.Right - Main.instance.xlAppRect.Left - (5 + left * 2);
                int height = (int)(120 * Main.instance.zoom / 100);

                menuForm.SetRects(left, top, width, height);

                if (Main.instance.xlApp.ActiveWindow.FreezePanes == false)
                {
                    Main.instance.xlApp.Range[range].Select();
                    Main.instance.xlApp.ActiveWindow.FreezePanes = false;
                    Main.instance.xlApp.ActiveWindow.FreezePanes = true;
                }
                Main.instance.xlApp.Range[oldRange].Select();
                Main.instance.xlApp.ActiveWindow.ScrollRow = scrollValueR;
                Main.instance.xlApp.ActiveWindow.ScrollColumn = scrollValueC;
            }
            Main.instance.ResumeAll();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            CalculateFormsPositions();
        }

        //// 
        //// tableLayoutPanel2
        //// 
        //this.tableLayoutPanel2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
        //this.tableLayoutPanel2.BackColor = System.Drawing.SystemColors.ActiveCaption;
        //this.tableLayoutPanel2.ColumnCount = 3;
        //this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
        //this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
        //this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
        //this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel5, 0, 0);
        //this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel4, 0, 0);
        //this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel3, 0, 0);
        //this.tableLayoutPanel2.Location = new System.Drawing.Point(1, 1);
        //this.tableLayoutPanel2.Margin = new System.Windows.Forms.Padding(1);
        //this.tableLayoutPanel2.Name = "tableLayoutPanel2";
        //this.tableLayoutPanel2.RightToLeft = System.Windows.Forms.RightToLeft.No;
        //this.tableLayoutPanel2.RowCount = 1;
        //this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle());
        //this.tableLayoutPanel2.Size = new System.Drawing.Size(444, 115);
        //this.tableLayoutPanel2.TabIndex = 54;
        //// 
        //// tableLayoutPanel5
        //// 
        //this.tableLayoutPanel5.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
        //| System.Windows.Forms.AnchorStyles.Left) 
        //| System.Windows.Forms.AnchorStyles.Right)));
        //this.tableLayoutPanel5.AutoSize = true;
        //this.tableLayoutPanel5.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
        //this.tableLayoutPanel5.ColumnCount = 1;
        //this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
        //this.tableLayoutPanel5.Controls.Add(this.button48);
        //this.tableLayoutPanel5.Controls.Add(this.checkBox1);
        //this.tableLayoutPanel5.Location = new System.Drawing.Point(334, 0);
        //this.tableLayoutPanel5.Margin = new System.Windows.Forms.Padding(0);
        //this.tableLayoutPanel5.Name = "tableLayoutPanel5";
        //this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle());
        //this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle());
        //this.tableLayoutPanel5.Size = new System.Drawing.Size(110, 115);
        //this.tableLayoutPanel5.TabIndex = 56;
        //// 
        //// tableLayoutPanel4
        //// 
        //this.tableLayoutPanel4.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
        //| System.Windows.Forms.AnchorStyles.Left) 
        //| System.Windows.Forms.AnchorStyles.Right)));
        //this.tableLayoutPanel4.AutoSize = true;
        //this.tableLayoutPanel4.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
        //this.tableLayoutPanel4.ColumnCount = 1;
        //this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
        //this.tableLayoutPanel4.Controls.Add(this.tbSearch);
        //this.tableLayoutPanel4.Controls.Add(this.listBox1);
        //this.tableLayoutPanel4.Location = new System.Drawing.Point(107, 0);
        //this.tableLayoutPanel4.Margin = new System.Windows.Forms.Padding(0);
        //this.tableLayoutPanel4.Name = "tableLayoutPanel4";
        //this.tableLayoutPanel4.RightToLeft = System.Windows.Forms.RightToLeft.No;
        //this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle());
        //this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle());
        //this.tableLayoutPanel4.Size = new System.Drawing.Size(227, 115);
        //this.tableLayoutPanel4.TabIndex = 54;
        //// 
        //// tableLayoutPanel3
        //// 
        //this.tableLayoutPanel3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
        //| System.Windows.Forms.AnchorStyles.Left) 
        //| System.Windows.Forms.AnchorStyles.Right)));
        //this.tableLayoutPanel3.AutoSize = true;
        //this.tableLayoutPanel3.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
        //this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
        //this.tableLayoutPanel3.Controls.Add(this.btnToArhive);
        //this.tableLayoutPanel3.Controls.Add(this.btnFromArhive);
        //this.tableLayoutPanel3.Controls.Add(this.button46);
        //this.tableLayoutPanel3.Controls.Add(this.button47);
        //this.tableLayoutPanel3.Location = new System.Drawing.Point(0, 0);
        //this.tableLayoutPanel3.Margin = new System.Windows.Forms.Padding(0);
        //this.tableLayoutPanel3.Name = "tableLayoutPanel3";
        //this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle());
        //this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle());
        //this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle());
        //this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle());
        //this.tableLayoutPanel3.Size = new System.Drawing.Size(107, 115);
        //this.tableLayoutPanel3.TabIndex = 54;
        //// 
        //// flowLayoutPanel2
        //// 
        //this.flowLayoutPanel2.AutoSize = true;
        //this.flowLayoutPanel2.Controls.Add(this.tableLayoutPanel2);
        //this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Right;
        //this.flowLayoutPanel2.Location = new System.Drawing.Point(1388, 0);
        //this.flowLayoutPanel2.Margin = new System.Windows.Forms.Padding(0);
        //this.flowLayoutPanel2.Name = "flowLayoutPanel2";
        //this.flowLayoutPanel2.Size = new System.Drawing.Size(446, 313);
        //this.flowLayoutPanel2.TabIndex = 56;
    }
}

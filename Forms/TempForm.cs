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
    }
}

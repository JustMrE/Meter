using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Text.RegularExpressions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace Meter
{
    public static class GlobalMethods
    {
        public static string logFile;
        public static string username;

        public static void CalculateFormsPositions()
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

                Main.instance.menu.SetRects(left, top, width, height);

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
    
        public static void ReleseObject(object o)
        {
            if (o is not null)
            {
                try
                {
                    Marshal.ReleaseComObject(o);
                }
                catch
                {
                    
                }
            }
        }

        public static void ToLog(string msg)
        {
            
            using (StreamWriter writer = new StreamWriter(logFile, true))
            {
                writer.WriteLine(DateTime.Now + " " + username + " " + msg);
            }
        }

        public static void ToLog(object sender)
        {
            ToLog("Нажато " + ((Control)sender).Name);
        }
        public static void ToLog(object sender, string txt)
        {
            ToLog("Изменен текст " + ((Control)sender).Name + " на '" + txt + "'");
        }
        public static void ToLog(object sender, object selectedItem)
        {
            ToLog("В списке " + ((Control)sender).Name + " выбран " + selectedItem.ToString());
        }

        public static void ToLog(Form form, object sender)
        {
            ToLog("Нажато " + ((Control)sender).Name + " на форме " + form.Name);
        }
        public static void ToLog(Form form, object sender, string txt)
        {
            ToLog("Изменен текст " + ((Control)sender).Name + " на '" + txt + "' на форме " + form.Name);
        }
        public static void ToLog(Form form, object sender, object selectedItem)
        {
            ToLog("В списке " + ((Control)sender).Name + " выбран " + selectedItem.ToString() + " на форме " + form.Name);
        }
        public static void ToLog(Form form, object sender, bool ch)
        {
            ToLog("Переключатель переключен на " + ((Control)sender).Name + " на '" + ch.ToString() + "' на форме " + form.Name);
        }
        public static void ToLog(Form form)
        {
            ToLog("Открыта форма " + form.Name);
        }

        public static void ClearLogs()
        {
            if (File.Exists(logFile))
            {
                using (File.CreateText(logFile));
            }
            ToLog("Логи очищены");
        }
    }
}

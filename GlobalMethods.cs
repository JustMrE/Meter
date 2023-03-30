using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Drawing;
using System.Text.RegularExpressions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace Meter
{
    public static class GlobalMethods
    {
        public static string logFile, errLogFile;
        public static string username;
        public static CultureInfo culture = new CultureInfo("ru-RU");
        public static float dpiX, dpiY;

        public static void CalculateFormsPositions()
        {
            //Main.instance.StopAll();
            //if (Main.instance.xlApp.WindowState != Excel.XlWindowState.xlMinimized && !Main.instance.closed)
            {
                Main.GetWindowRect(Main.instance.xlAppHwnd, ref Main.instance.xlAppRect);
                Main.instance.xlAppRect.Top = Main.instance.xlAppRect.Top < 0 ? 0 : Main.instance.xlAppRect.Top;
                Main.instance.xlAppRect.Left = Main.instance.xlAppRect.Left < 0 ? 0 : Main.instance.xlAppRect.Left;
                
                int yCorrect = Convert.ToInt32((double)((Excel.Range)Main.instance.xlApp.ActiveWindow.ActivePane.VisibleRange.Cells[1,1]).Top - (double)((Excel.Worksheet)Main.instance.wb.ActiveSheet).Range["B7"].Top);
                int xCorrect = Convert.ToInt32((double)((Excel.Range)Main.instance.xlApp.ActiveWindow.ActivePane.VisibleRange.Cells[1,1]).Left - (double)((Excel.Worksheet)Main.instance.wb.ActiveSheet).Range["B7"].Left);

                int left = Main.instance.xlApp.ActiveWindow.ActivePane.PointsToScreenPixelsX(0 + xCorrect) - Main.instance.xlAppRect.Left;
                int top = Main.instance.xlApp.ActiveWindow.ActivePane.PointsToScreenPixelsY(0 + yCorrect) - Main.instance.xlAppRect.Top;

                Excel.Range visibleRange = Main.instance.xlApp.ActiveWindow.ActivePane.VisibleRange;
                double rangeSizeX = ((double)visibleRange.Width / 72f) * dpiX;
                int width = (int)(rangeSizeX * (double)Main.instance.xlApp.ActiveWindow.Zoom / 100);
                
                double rangeSizeY = ((double)((Excel.Worksheet)Main.instance.wb.ActiveSheet).Range["B1:B6"].Height / 72f) * dpiY;
                int height = (int)(rangeSizeY * (double)Main.instance.xlApp.ActiveWindow.Zoom / 100);
                
                Main.instance.menu.SetRects(left, top, width, height);
            }
            //Main.instance.ResumeAll();
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
        
        public static void Err(string msg)
        {
            using (StreamWriter writer = new StreamWriter(errLogFile, true))
            {
                writer.WriteLine(DateTime.Now + " " + username + " " + msg);
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
            if (File.Exists(errLogFile))
            {
                using (File.CreateText(errLogFile));
            }
            ToLog("Логи очищены");
        }
    }
}

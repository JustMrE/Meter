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
using System.Diagnostics;
using System.Runtime.Serialization;
// using AArch;

namespace Meter
{
    public static class AArchTest 
    {
        static readonly Guid AArchGuid = new Guid("089BE542-9C7B-11D1-8419-006094AC7BF7");

        public static void ReadTI (Int32 TI, DateTime date)
        {
            // string date = Console.ReadLine();
            // string format = "dd.MM.yyyy HH:mm:ss";
            // DateTime Текущее_Время = DateTime.Now;

            // try
            // {
            //     Текущее_Время = DateTime.ParseExact(date, format, null);
            // }
            // catch (FormatException e)
            // {
            //     Console.WriteLine("Not correct date");
            //     return;
            // }

            try
            {
                dynamic AutoArch = Activator.CreateInstance(Type.GetTypeFromCLSID(AArchGuid));
                AutoArch.SetBanks();
                AutoArch.SetBase();

                DateTime Текущее_Время = date;
                Int32 РК_ТИ = TI;

                Double Генерация_РК_ОИК = (Double)AutoArch.GetTI(Текущее_Время, РК_ТИ);
                MessageBox.Show(String.Format("{0:0.00}", Генерация_РК_ОИК));
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return;
            }
            
        }
        
    }
}
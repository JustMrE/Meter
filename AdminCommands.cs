using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Meter.Forms;
using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;

namespace Meter 
{
    public static class AdminCommands
    {
        static Dictionary<string, Action> commands = new ();

        public static void SetupCommands()
        {
            commands.Add("dontsave", () => 
            {
                MeterSettings.Instance.CloseAutoSave = false;
            });
            commands.Add("save", () => 
            {
                MeterSettings.Instance.CloseAutoSave = true;
            });
            commands.Add("stop", () => 
            {
                Main.instance.xlApp.EnableEvents = false;
            });
            commands.Add("resume", () => 
            {
                Main.instance.xlApp.EnableEvents = true;
            });
        }

        public static void RunCommand(string cmd)
        {
            try
            {
                if (commands.ContainsKey(cmd))
                {
                    GlobalMethods.ToLogWarn($"Выполнена команда администратора: `{cmd}`" );
                    commands[cmd].Invoke();
                }
            }
            catch (System.Exception)
            {
                
            }
        }

    }
}
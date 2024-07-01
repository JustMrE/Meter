using System.Diagnostics;
using Meter.Forms;

namespace Meter
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        /*[STAThread]
        static void Main()
        {
            string? dir = null;
            string? db = null;
            bool pathExists = false;

            string openedFlagFile;
            string logFile;
            string errLogFile;
            string username;

            string dbPathInfoFile = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\db.txt";
            //// To customize application configuration such as set high DPI settings or default font,
            //// see https://aka.ms/applicationconfiguration.
            ////ApplicationConfiguration.Initialize();
            ////Application.Run(new Menu());
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            if (File.Exists(dbPathInfoFile))
            {
                dir = File.ReadAllText(dbPathInfoFile);
                if (Directory.Exists(dir))
                {
                    db = dir;
                    pathExists = true;
                }
            }

            if (!pathExists)
            {
                string? selectionPath = null;
                OpenFileDialog folderBrowserDialog = new()
                {
                    Title = "Выберите запускаемый файл счетчиков (Meter.exe)"
                };

                DialogResult result = folderBrowserDialog.ShowDialog();

                if (result == DialogResult.OK)
                {
                    selectionPath = folderBrowserDialog.FileName;
                }
                else
                {
                    return;
                }
                dir = selectionPath;
                db = null;
                if (!string.IsNullOrEmpty(dir))
                {
                    db = Path.GetDirectoryName(dir) + @"\DB";
                    if (!Directory.Exists(db))
                    {
                        MessageBox.Show("База данных не найдена!");
                        return;
                    }
                }

                string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\db.txt";
                using (StreamWriter writer = File.CreateText(path))
                {
                    writer.Write(db);
                    File.SetAttributes(db, File.GetAttributes(db) | FileAttributes.Hidden);
                }
            }

            openedFlagFile = db + @"\opened.txt";
            logFile = db + @"\current\log.log";
            errLogFile = db + @"\current\errlog.log";
            username = Environment.UserName;

            #if !DEBUG
            if (File.Exists(openedFlagFile))
            {
                string newUserName = File.ReadAllText(openedFlagFile);
                if (!string.IsNullOrEmpty(newUserName))
                {
                    MessageBox.Show("Счетчики уже открыты пользователем " + newUserName);
                    return;
                }
            }
            #endif

            GlobalMethods.username = username;
            GlobalMethods.logFile = logFile;
            GlobalMethods.errLogFile = errLogFile;

            #if !DEBUG
            using (StreamWriter writer = File.CreateText(openedFlagFile))
            {
                writer.Write(username);
                File.SetAttributes(openedFlagFile, File.GetAttributes(openedFlagFile) | FileAttributes.Hidden);
            }
            #endif

            using (StreamWriter writer = new (logFile, true)) 
            {
                writer.WriteLine();
                writer.WriteLine();
                writer.WriteLine(DateTime.Now +  " Счетчики открыты пользователем " + username);
            }
            using (StreamWriter writer = new (errLogFile, true)) 
            {
                writer.WriteLine();
                writer.WriteLine();
                writer.WriteLine(DateTime.Now +  " Счетчики открыты пользователем " + username);
            }


            MyApplicationContext myAppContext = new ();
            Application.Run(myAppContext);

            #if !DEBUG
            File.Delete(openedFlagFile);
            #endif
        }*/

        [STAThread]
        static void Main()
        {
            GlobalMethods.username = Environment.UserName;
            string openedFlagFile = string.Empty;

            //// To customize application configuration such as set high DPI settings or default font,
            //// see https://aka.ms/applicationconfiguration.
            ////ApplicationConfiguration.Initialize();
            ////Application.Run(new Menu());
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // new MeterSettings(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\settings");
            MeterSettings.Instance.SettingsFile = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\settings";
            if (!File.Exists(MeterSettings.Instance.SettingsFile) || !MeterSettings.Instance.Load())
            {
                Setup settings = new Setup();
                DialogResult result = settings.ShowDialog();
                if (result != DialogResult.OK)
                {
                    return;
                }
            }

            #if !DEBUG
            openedFlagFile = MeterSettings.Instance.DBDir + @"\opened";
            if (File.Exists(openedFlagFile))
            {
                string newUserName = File.ReadAllText(openedFlagFile);
                if (!string.IsNullOrEmpty(newUserName))
                {
                    MessageBox.Show("Счетчики уже открыты пользователем " + newUserName);
                    return;
                }
            }
            using (StreamWriter writer = File.CreateText(openedFlagFile))
            {
                writer.Write(GlobalMethods.username);
                File.SetAttributes(openedFlagFile, File.GetAttributes(openedFlagFile) | FileAttributes.Hidden);
            }
            #endif

            using (StreamWriter writer = new (MeterSettings.Instance.LogFile, true)) 
            {
                writer.WriteLine();
                writer.WriteLine();
                writer.WriteLine(DateTime.Now +  " Счетчики открыты пользователем " + GlobalMethods.username);
            }
            using (StreamWriter writer = new (MeterSettings.Instance.ErrLogFile, true)) 
            {
                writer.WriteLine();
                writer.WriteLine();
                writer.WriteLine(DateTime.Now +  " Счетчики открыты пользователем " + GlobalMethods.username);
            }


            MyApplicationContext myAppContext = new ();
            Application.Run(myAppContext);

            #if !DEBUG
            File.Delete(openedFlagFile);
            #endif
        }
    }
}
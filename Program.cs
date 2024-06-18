using System.Diagnostics;

namespace Meter
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //// To customize application configuration such as set high DPI settings or default font,
            //// see https://aka.ms/applicationconfiguration.
            ////ApplicationConfiguration.Initialize();
            ////Application.Run(new Menu());
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string? dir = null;
            string? db = null;
            bool pathExists = false;

            string f = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\db.txt";
            if (File.Exists(f))
            {
                dir = File.ReadAllText(f);
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

                // MessageBox.Show(dir);

                if (!string.IsNullOrEmpty(dir))
                {
                    db = Path.GetDirectoryName(dir) + @"\DB";
                    // MessageBox.Show(db);
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

            string file1 = db + @"\opened.txt";
            string logFile = db + @"\current\log.log";
            string errLogFile = db + @"\current\errlog.log";
            string username = Environment.UserName;
            #if !DEBUG
            if (File.Exists(file1))
            {
                string newUserName = File.ReadAllText(file1);
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
            using (StreamWriter writer = File.CreateText(file1))
            {
                writer.Write(username);
                File.SetAttributes(file1, File.GetAttributes(file1) | FileAttributes.Hidden);
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
            File.Delete(file1);
            #endif
        }    
    }
}
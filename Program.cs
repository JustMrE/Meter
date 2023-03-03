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
            string dir = Process.GetCurrentProcess().MainModule.FileName;
            string file = System.IO.Path.GetDirectoryName(dir) + @"\Счетчики.xlsm";
            if (!File.Exists(file))
            {
                dir = System.Reflection.Assembly.GetExecutingAssembly().Location;
                file = System.IO.Path.GetDirectoryName(dir) + @"\Счетчики.xlsm";
            }
            
            //string file1 = System.IO.Path.GetDirectoryName(dir) + @"\DB\opened.txt";
            //string username = Environment.UserName;
            //if (File.Exists(file1))
            //{
            //    string newUserName = File.ReadAllText(file1);
            //    if (!string.IsNullOrEmpty(newUserName))
            //    {
            //        MessageBox.Show("Счетчики уже открыты пользователем " + newUserName);
            //        return;
            //    }
            //}
            //using (StreamWriter writer = File.CreateText(file1))
            //{
            //    writer.Write(username);
            //    File.SetAttributes(file1, File.GetAttributes(file1) | FileAttributes.Hidden);
            //}

            //// To customize application configuration such as set high DPI settings or default font,
            //// see https://aka.ms/applicationconfiguration.
            ////ApplicationConfiguration.Initialize();
            ////Application.Run(new Menu());
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            MyApplicationContext myAppContext = new MyApplicationContext();
            Application.Run(myAppContext);
            //File.Delete(file1);
        }    
    }
}
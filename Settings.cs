

using Newtonsoft.Json;

namespace Meter.Forms
{
    public static class MeterSettings
    {
        public static string settingsFile = string.Empty;
        static string meterFile = string.Empty;
        static string dbDir = string.Empty;
        static string logFile = string.Empty;
        static string errLogFile = string.Empty;
        static bool closeAutoSave = false;
        public static string DBDir
        {
            get 
            {
                return dbDir;
            }
            set
            {
                dbDir = value;
            }
        }
        public static string MeterFile
        {
            get
            {
                return meterFile;
            }
            set
            {
                meterFile = value;
            }
        }
        public static string LogFile
        {
            get 
            {
                return logFile;
            }
            set
            {
                logFile = value;
            }
        }
        public static string ErrLogFile
        {
            get
            {
                return errLogFile;
            }
            set
            {
                errLogFile = value;
            }
        }
        public static bool CloseAutoSave
        {
            get;
            set;
        }
        public static bool Load()
        {
            if (!string.IsNullOrEmpty(settingsFile) && File.Exists(settingsFile))
            {
                using (StreamReader sr = new StreamReader(settingsFile))
                {
                    var settingsJson = sr.ReadToEnd();
                    MeterSettingsSerialization settings = JsonConvert.DeserializeObject<MeterSettingsSerialization>(settingsJson);
                    if (settings != null)
                    {
                        meterFile = settings.meterFile;
                        dbDir = settings.dbDir;
                        logFile = settings.logFile;
                        errLogFile = settings.errLogFile;
                        closeAutoSave = closeAutoSave;

                        return true;
                    }
                    else
                    {
                        MessageBox.Show("Ошибка при загрузке настроек!");
                        return false;
                    }
                }
            }
            else
            {
                string pathToMeter;
                OpenFileDialog folderBrowserDialog = new()
                {
                    Title = "Выберите запускаемый файл счетчиков (Meter.exe)"
                };
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    pathToMeter = folderBrowserDialog.FileName;
                }
                else
                {
                    return false;
                }
                dbDir = pathToMeter;
                if (!string.IsNullOrEmpty(dbDir))
                {
                    dbDir = Path.GetDirectoryName(dbDir) + @"\DB";
                    if (!Directory.Exists(dbDir))
                    {
                        MessageBox.Show("База данных не найдена!");
                        return false;
                    }
                    else
                    {
                        MeterFile = dbDir + @"\current\meter.xlsx";
                        logFile = dbDir + @"\current\log.log";
                        errLogFile = dbDir + @"\current\errlog.log";
                        closeAutoSave = false;

                        Save();
                        return true;
                    }
                }
                // string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\db.txt";
                // using (StreamWriter writer = File.CreateText(path))
                // {
                //     writer.Write(db);
                //     File.SetAttributes(db, File.GetAttributes(db) | FileAttributes.Hidden);
                // }
            }
            return false;
        }

        public static void Save()
        {
            // string file = MeterSettings.DBDir + path + @"\colors.json";
            using (StreamWriter writer = File.CreateText(settingsFile))
            {
                MeterSettingsSerialization settings = new () {
                    meterFile = meterFile,
                    dbDir = dbDir,
                    logFile = logFile,
                    errLogFile = errLogFile,
                    closeAutoSave = closeAutoSave,
                };
                var jsonString = JsonConvert.SerializeObject(settings);
                writer.Write(jsonString);
            }
        }
        class MeterSettingsSerialization
        {
            public string meterFile { get; set; }
            public string dbDir { get; set; }
            public string logFile { get; set; }
            public string errLogFile { get; set; }
            public bool closeAutoSave { get; set; }
        }
    }
}


using System.Reflection;
using Newtonsoft.Json;

namespace Meter.Forms
{
    public class MeterSettings
    {
        static MeterSettings _Instance = new MeterSettings();

        public string SettingsFile { get; set; }
        public string DBDir { get; set; }
        public string MeterFile { get; set; }
        public string LogFile { get; set; }
        public string ErrLogFile { get; set; }
        public bool CloseAutoSave { get; set; }

        public static MeterSettings Instance 
        {
            get { return _Instance; }
        }

        public MeterSettings ()
        {
            SettingsFile = string.Empty;
            DBDir = string.Empty;
            MeterFile = string.Empty;
            LogFile = string.Empty;
            ErrLogFile = string.Empty;
            CloseAutoSave = false;
        }

        public bool Load()
        {
            try
            {
                var jsonString = File.ReadAllText(SettingsFile);
                var loadedSettings = JsonConvert.DeserializeObject<MeterSettings>(jsonString);

                // Обновление текущего экземпляра свойствами из десериализованного объекта
                if (loadedSettings != null)
                {
                    CopyPropertiesFrom(loadedSettings);
                }
                if (!CheckDBFiles())
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
        
        public void Save()
        {
            try
            {
                using (StreamWriter writer = File.CreateText(SettingsFile))
                {
                    var jsonString = JsonConvert.SerializeObject(this);
                    writer.Write(jsonString);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving settings: {ex.Message}");
            }
        }

        private void CopyPropertiesFrom(MeterSettings other)
        {
            foreach (PropertyInfo property in typeof(MeterSettings).GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (property.CanRead && property.CanWrite)
                {
                    property.SetValue(this, property.GetValue(other, null), null);
                }
            }
        }

        public bool CheckDBFiles()
        {
            string arch = DBDir + @"\arch";

            string current = DBDir + @"\current";
            string formulas = current + @"\formulas";
            string references = current + @"\references";
            string tiDictFile = current + @"\Словарь ТИ факт.xlsx";
            string colors = current + @"\colors.json";

            string saves = DBDir + @"\saves";
            string savedFormulas = saves + @"\formulas";
            string tempArch = DBDir + @"\temparch";

            string standartColors = DBDir + @"\standartColors.json";

            bool hasPaths = !string.IsNullOrEmpty(DBDir) && !string.IsNullOrEmpty(MeterFile) && !string.IsNullOrEmpty(LogFile) && !string.IsNullOrEmpty(ErrLogFile);
            bool hasDirs = Directory.Exists(DBDir) && Directory.Exists(arch) && Directory.Exists(current) && Directory.Exists(formulas) && Directory.Exists(references) && Directory.Exists(saves) && Directory.Exists(savedFormulas) && Directory.Exists(tempArch);
            bool hasFiles = File.Exists(MeterFile) && File.Exists(tiDictFile) && File.Exists(colors) && File.Exists(standartColors);
            return hasPaths && hasDirs && hasFiles;
        }
    }
}
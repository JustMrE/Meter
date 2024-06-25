using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Main = Meter.MyApplicationContext;
using Meter.Forms;

namespace Meter
{
    public static class SaveLoader
    {
        public static void SaveAsync()
        {
            GlobalMethods.ToLog("Сохранение данных в папку source");
            var tasks = new List<Task>();

            var task = Task.Run(() => SaveReferencesNew());
            tasks.Add(task);

            task = Task.Run(() => SaveColors());
            tasks.Add(task);

            task = Task.Run(() => SaveHeads());
            tasks.Add(task);

            Task.WaitAll(tasks.ToArray());
            GlobalMethods.ToLog("Данные сохранены в папку source");
        }

        public async static void LoadAsync()
        {
            GlobalMethods.ToLog("Загрузка данных из папки source");
            Main.loading = true;
            var tasks = new List<Task>();

            var task = Task.Run(() => LoadReferencesNew());
            tasks.Add(task);

            task = Task.Run(() => LoadColors());
            tasks.Add(task);

            task = Task.Run(() => LoadHeads());
            tasks.Add(task);

            Task.WaitAll(tasks.ToArray());
            Main.loading = false;
            GlobalMethods.ToLog("Данные загружены из папки source");
        }

        public async static void LoadAsyncFromFolder(string folder)
        {
            GlobalMethods.ToLog("Загрузка данных из папки temp");
            Main.loading = true;
            string path = @"\" + folder;
            LoadReferencesNew(path);
            LoadColors(path);
            LoadHeads(path);
            Main.loading = false;
            GlobalMethods.ToLog("Данные загружены из папки temp");
        }
        
        static async void SaveReferencesNew(string path = @"\current")
        {
            var tasks = new List<Task>();
            Directory.Delete(MeterSettings.Instance.DBDir + path + @"\references", true);
            Directory.CreateDirectory(MeterSettings.Instance.DBDir + path + @"\references");
            Directory.Delete(MeterSettings.Instance.DBDir + path + @"\formulas", true);
            Directory.CreateDirectory(MeterSettings.Instance.DBDir + path + @"\formulas");

            foreach (string n in Main.instance.references.references.Keys)
            {
                var task = Task.Run(() => 
                {
                    string name = Main.instance.references.references[n].ID;
                    string file = MeterSettings.Instance.DBDir  + path + @"\references\" + name + ".json";
                    using (StreamWriter writer = File.CreateText(file))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        var jsonString = JsonConvert.SerializeObject(Main.instance.references[n]);
                        writer.Write(jsonString);
                    }
                });
                tasks.Add(task);
            }

            foreach (string id in Main.instance.formulas.formulas.Keys)
            {
                var task = Task.Run(() => 
                {
                    string name = id;
                    string file = MeterSettings.Instance.DBDir + path + @"\formulas\" + name + ".json";
                    using (StreamWriter writer = File.CreateText(file))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        var jsonString = JsonConvert.SerializeObject(Main.instance.formulas.formulas[id]);
                        writer.Write(jsonString);
                    }
                });
                tasks.Add(task);
            }
            Task.WaitAll(tasks.ToArray());
        }
        
        static void LoadReferencesNew(string path = @"\current")
        {
            var tasks = new List<Task>();
            string path1 = MeterSettings.Instance.DBDir + path + @"\references";
            List<string> filePaths1 = Directory.GetFiles(path1, "*.json").ToList();
            ConcurrentDictionary<string, ReferenceObject> concurentDictionary1 = new ConcurrentDictionary<string, ReferenceObject>();
            foreach (var file in filePaths1)
            {
                var task = Task.Run(() => 
                {
                    using (var streamReader = new StreamReader(file))
                    {
                        var json = streamReader.ReadToEnd();
                        ReferenceObject ro = JsonConvert.DeserializeObject<ReferenceObject>(json);
                        string name = Path.GetFileName(file).Replace(".json","");
                        concurentDictionary1.TryAdd(ro._name, ro);
                    }
                });
                tasks.Add(task);
            }

            string path2 = MeterSettings.Instance.DBDir + path + @"\formulas";
            List<string> filePaths2 = Directory.GetFiles(path2, "*.json").ToList();
            ConcurrentDictionary<string, List<ForTags>> concurentDictionary2 = new ConcurrentDictionary<string, List<ForTags>>();
            foreach (var file in filePaths2)
            {
                var task = Task.Run(() => 
                {
                    using (var streamReader = new StreamReader(file))
                    {
                        var json = streamReader.ReadToEnd();
                        var formula = JsonConvert.DeserializeObject<List<ForTags>>(json);
                        string name = Path.GetFileName(file).Replace(".json","");
                        concurentDictionary2.TryAdd(name, formula);
                    }
                });
                tasks.Add(task);
            }

            Task.WaitAll(tasks.ToArray());

            if (concurentDictionary1.Count != 0)
            {
                Main.instance.references.references = concurentDictionary1.ToDictionary(kv => kv.Key, kv => kv.Value);
            }
            else
            {
                Main.instance.references.references = new Dictionary<string, ReferenceObject>();
            }

            if (concurentDictionary2.Count != 0)
            {
                Main.instance.formulas.formulas = concurentDictionary2.ToDictionary(kv => kv.Key, kv => kv.Value);
            }
            else
            {
                Main.instance.references.formulas = new Dictionary<string, List<ForTags>>();
            }

            Main.instance.references.UpdateAllLevels();
            Main.instance.references.UpdateAllParents();
            Main.instance.heads.UpdateParents();
        }

        static void SaveColors(string path = @"\current")
        {
            string file = MeterSettings.Instance.DBDir + path + @"\colors.json";
            using (StreamWriter writer = File.CreateText(file))
            {
                JsonSerializer serializer = new JsonSerializer();
                var jsonString = JsonConvert.SerializeObject(Main.instance.colors);
                writer.Write(jsonString);
            }
        }
        static void LoadColors(string path = @"\current")
        {
            string file = MeterSettings.Instance.DBDir + path + @"\colors.json";
            var stringJson = File.ReadAllText(file);
            Main.instance.colors = JsonConvert.DeserializeObject<ColorsData>(stringJson);
        }

        static void SaveHeads(string path = @"\current")
        {
            string file = MeterSettings.Instance.DBDir + path + @"\heads.json";
            using (StreamWriter writer = File.CreateText(file))
            {
                JsonSerializer serializer = new JsonSerializer();
                var jsonString = JsonConvert.SerializeObject(Main.instance.heads);
                writer.Write(jsonString);
            }
        }
        static void LoadHeads(string path = @"\current")
        {
            string file = MeterSettings.Instance.DBDir + path + @"\heads.json";
            if (File.Exists(file))
            {
                var stringJson = File.ReadAllText(file);
                HeadReferences hr = JsonConvert.DeserializeObject<HeadReferences>(stringJson);
                if (hr != null)
                {
                    Main.instance.heads = hr;
                }
                else
                {
                    Main.instance.heads = new HeadReferences();
                }
            }
            else
            {
                Main.instance.heads = new HeadReferences();
            }
            
        }

        public static void LoadStandartColors()
        {
            string file = MeterSettings.Instance.DBDir + @"\standartColors.json";
            var stringJson = File.ReadAllText(file);
            //var root = JsonConvert.DeserializeObject(stringJson).ToString();
            Main.instance.colors = JsonConvert.DeserializeObject<ColorsData>(stringJson);
        }
    }
}

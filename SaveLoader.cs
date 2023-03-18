using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Main = Meter.MyApplicationContext;

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
        
        static async void SaveReferencesNew()
        {
            var tasks = new List<Task>();
            foreach (string n in Main.instance.references.references.Keys)
            {
                var task = Task.Run(() => 
                {
                    string name = Main.instance.references.references[n].ID;
                    string file = Main.dir + @"\current\references\" + name + ".json";
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
                    string file = Main.dir + @"\current\formulas\" + name + ".json";
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
            //await Task.WhenAll(tasks);
        }
        
        static void LoadReferencesNew(string path = @"\current")
        {
            var tasks = new List<Task>();
            string path1 = Main.dir + path + @"\references";
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

            string path2 = Main.dir + path + @"\formulas";
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

            Main.instance.references.references = concurentDictionary1.ToDictionary(kv => kv.Key, kv => kv.Value);
            Main.instance.formulas.formulas = concurentDictionary2.ToDictionary(kv => kv.Key, kv => kv.Value);

            Main.instance.references.UpdateAllLevels();
            Main.instance.references.UpdateAllParents();
            Main.instance.heads.UpdateParents();
            Main.instance.heads.UpdateIndents(false);
        }

        static void SaveColors()
        {
            string file = Main.dir + @"\current\colors.json";
            using (StreamWriter writer = File.CreateText(file))
            {
                JsonSerializer serializer = new JsonSerializer();
                var jsonString = JsonConvert.SerializeObject(Main.instance.colors);
                //serializer.Serialize(writer, jsonString);
                writer.Write(jsonString);
            }
        }
        static void LoadColors(string path = @"\current")
        {
            string file = Main.dir + path + @"\colors.json";
            var stringJson = File.ReadAllText(file);
            //var root = JsonConvert.DeserializeObject(stringJson).ToString();
            Main.instance.colors = JsonConvert.DeserializeObject<ColorsData>(stringJson);
        }

        static void SaveHeads()
        {
            string file = Main.dir + @"\current\heads.json";
            using (StreamWriter writer = File.CreateText(file))
            {
                JsonSerializer serializer = new JsonSerializer();
                var jsonString = JsonConvert.SerializeObject(Main.instance.heads);
                writer.Write(jsonString);
            }
        }
        static void LoadHeads(string path = @"\current")
        {
            string file = Main.dir + path + @"\heads.json";
            var stringJson = File.ReadAllText(file);
            Main.instance.heads = JsonConvert.DeserializeObject<HeadReferences>(stringJson);
        }

        public static void LoadStandartColors()
        {
            string file = Main.dir + @"\standartColors.json";
            var stringJson = File.ReadAllText(file);
            //var root = JsonConvert.DeserializeObject(stringJson).ToString();
            Main.instance.colors = JsonConvert.DeserializeObject<ColorsData>(stringJson);
        }
    }
}

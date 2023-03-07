using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;

namespace Meter
{
    public static class SaveLoader
    {
        public static void Save()
        {
            SaveReferences();
            SaveColors();
            SaveHeads();
        }

        public static void SaveAsync()
        {
            SaveReferencesNew();
            SaveColors();
            SaveHeads();
        }

        public async static void Load()
        {
            LoadReferences();
            LoadColors();
            LoadHeads();
        }

        public async static void LoadAsync()
        {
            LoadReferencesNew();
            LoadColors();
            LoadHeads();
        }

        static void SaveReferences()
        {
            string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\referencesDictionary.json";
            using (StreamWriter writer = File.CreateText(file))
            {
                JsonSerializer serializer = new JsonSerializer();
                var jsonString = JsonConvert.SerializeObject(Main.instance.references);
                writer.Write(jsonString);
                //serializer.Serialize(writer, jsonString);
            }

            file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\formulasDictionary.json";
            using (StreamWriter writer = File.CreateText(file))
            {
                JsonSerializer serializer = new JsonSerializer();
                var jsonString = JsonConvert.SerializeObject(Main.instance.formulas);
                writer.Write(jsonString);
                //serializer.Serialize(writer, jsonString);
            }
        }
        static async void SaveReferencesNew()
        {
            var tasks = new List<Task>();
            foreach (string n in Main.instance.references.references.Keys)
            {
                var task = Task.Run(() => 
                {
                    string name = Main.instance.references.references[n].ID;
                    string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\references\" + name + ".json";
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
                    string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\formulas\" + name + ".json";
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
        static void LoadReferences()
        {
            string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\referencesDictionary.json";
            var stringJson = File.ReadAllText(file);
            //var root = JsonConvert.DeserializeObject(stringJson).ToString();
            Main.instance.references = JsonConvert.DeserializeObject<RangeReferences>(stringJson);

            Main.instance.references.UpdateAllLevels();
            Main.instance.references.UpdateAllParents();

            file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\formulasDictionary.json";
            if (File.Exists(file))
            {
                stringJson = File.ReadAllText(file);
                //root = JsonConvert.DeserializeObject(stringJson).ToString();
                Main.instance.formulas = JsonConvert.DeserializeObject<Formula>(stringJson);
            }
        }
        static void LoadReferencesNew()
        {
            var tasks = new List<Task>();
            string path1 = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\references";
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

            string path2 = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\formulas";
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
        }

        static void SaveColors()
        {
            string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\colors.json";
            using (StreamWriter writer = File.CreateText(file))
            {
                JsonSerializer serializer = new JsonSerializer();
                var jsonString = JsonConvert.SerializeObject(Main.instance.colors);
                //serializer.Serialize(writer, jsonString);
                writer.Write(jsonString);
            }
        }
        static void LoadColors()
        {
            string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\colors.json";
            var stringJson = File.ReadAllText(file);
            //var root = JsonConvert.DeserializeObject(stringJson).ToString();
            Main.instance.colors = JsonConvert.DeserializeObject<ColorsData>(stringJson);
        }

        static void SaveHeads()
        {
            string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\heads.json";
            using (StreamWriter writer = File.CreateText(file))
            {
                JsonSerializer serializer = new JsonSerializer();
                var jsonString = JsonConvert.SerializeObject(Main.instance.heads);
                writer.Write(jsonString);
            }
        }
        static void LoadHeads()
        {
            string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\heads.json";
            var stringJson = File.ReadAllText(file);
            Main.instance.heads = JsonConvert.DeserializeObject<HeadReferences>(stringJson);
        }

        public static void LoadStandartColors()
        {
            string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\colors.json";
            var stringJson = File.ReadAllText(file);
            //var root = JsonConvert.DeserializeObject(stringJson).ToString();
            Main.instance.colors = JsonConvert.DeserializeObject<ColorsData>(stringJson);
        }
    }
}

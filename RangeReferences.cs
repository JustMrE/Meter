﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Main = Meter.MyApplicationContext;
using Microsoft.Office.Interop.Excel;
using System.Timers;
using System.Diagnostics;
using Meter.Forms;

namespace Meter
{
    public class RangeReferences
    {
        [JsonIgnore]
        public List<string> Errors;
        [JsonIgnore]
        public Dictionary<string, List<ForTags>> formulas;
        //public List<ReferenceObject> references { get; set; }
        public Dictionary<string, ReferenceObject> references {get; set; }

        [JsonIgnore]
        public static ReferenceObject activeTable;
        [JsonIgnore]
        public static string _range;
        [JsonIgnore]
        public static ChildObject _activeObject;
        [JsonIgnore]
        public static List<ChildObject> IDErrors1 = new List<ChildObject>(), IDErrors2 = new List<ChildObject>();

        public RangeReferences()
        {
            //references = new List<ReferenceObject>();
            references = new Dictionary<string, ReferenceObject>();
        }

        [JsonIgnore]
        public static Dictionary<string, ReferencesParent> idDictionary = new Dictionary<string, ReferencesParent>();

        [JsonIgnore]
        public ReferenceObject this[string name]
        {
            get
            {
                //return references.Where(n => n._name == name).FirstOrDefault();
                return references[name];
            }
        }

        [JsonIgnore]
        public ReferenceObject this[Excel.Range rng]
        {
            get
            {
                //return references.Values.Where(n => n.HasRange(rng)).FirstOrDefault();
                return references.Values.AsParallel().FirstOrDefault(n => n.HasRange(rng));
            }
        }

        [JsonIgnore]
        public static string? ActiveL1
        {
            get
            {
                return activeTable.PS._activeChild != null ? activeTable.PS._activeChild._name : null;
            }
        }

        [JsonIgnore]
        public static string? activeL2
        {
            get
            {
                return activeTable.PS._activeChild._activeChild != null ? activeTable.PS._activeChild._activeChild._name : null;
            }
        }

        public void ActivateTable(Excel.Range rng)
        {
            activeTable = null;
            activeTable = this[rng];
            if (activeTable != null)
            {
                activeTable.ActivateTable(rng);
            }
        }

        public void CreateNew(string name, string nameL1, string psAddress, bool insert = true, bool stopall = true)
        {
            GlobalMethods.ToLog("Добавлен субъект {" + name + "}");
            ReferenceObject ro = new ReferenceObject(name, nameL1, psAddress, insert, stopall);
            references.Add(name, ro);
        }

        public bool Contains(string name)
        {
            //return references.Where(n => n._name == name).FirstOrDefault() != null;
            return references.ContainsKey(name);
        }
        public void UpdateAllColors()
        {
            foreach (ReferenceObject item in references.Values)
            {
                item.UpdateAllColors();
            }
            MessageBox.Show("Done!");
        }
        public void UpdateAllColors1(bool stopall = true)
        {
            if (stopall) Main.instance.StopAll();
            foreach (ReferenceObject item in references.Values)
            {
                item.UpdateAllColors();
            }
            if (stopall) Main.instance.ResumeAll();
        }
        public void UpdateAllBorders()
        {
            Main.instance.StopAll();
            foreach (ReferenceObject item in references.Values)
            {
                item.UpdateBorders();
            }
            Main.instance.ResumeAll();
        }
        public void UpdateAllNames()
        {
            foreach (ReferenceObject item in references.Values)
            {
                item.UpdateNames();
            }
            MessageBox.Show("Done!");
        }
        public void UpdateAllDBNames()
        {
            var watch = Stopwatch.StartNew();
            Main.instance.StopAll();
            foreach (ReferenceObject item in references.Values)
            {
                item.UpdateDBNames();
            }
            Main.instance.ResumeAll();
            watch.Stop();
            MessageBox.Show("Done!\n" + (watch.ElapsedMilliseconds / 1000) + " sec");
        }
        public void UpdateAllPSFormulas()
        {
            var watch = Stopwatch.StartNew();
            Main.instance.StopAll();
            foreach (ReferenceObject item in references.Values)
            {
                item.UpdateAllPSFormulas(false);
            }
            Main.instance.ResumeAll();
            watch.Stop();
            MessageBox.Show("Done!\n" + (watch.ElapsedMilliseconds / 1000) + " sec");
        }

        public void UpdateAllDBFormulas()
        {
            var watch = Stopwatch.StartNew();
            Main.instance.StopAll();
            foreach (ReferenceObject item in references.Values)
            {
                GlobalMethods.ToLog(item._name);
                item.UpdateAllDBFormulas(false);
            }
            Main.instance.ResumeAll();
            watch.Stop();
            MessageBox.Show("Done!\n" + (watch.ElapsedMilliseconds / 1000) + " sec");
        }
        public void UpdateAllParents(bool message = false)
        {
            idDictionary = new Dictionary<string, ReferencesParent>();
            foreach (ReferenceObject item in references.Values)
            {
                item.UpdateParents();
            }
            if (message) MessageBox.Show("Done!");
        }
        public void UpdateAllReferencesPS()
        {
            var watch = Stopwatch.StartNew();
            Main.instance.StopAll();
            foreach (ReferenceObject item in references.Values)
            {
                item.UpdateReferencesPS();
            }
            Main.instance.ResumeAll();
            watch.Stop();
            MessageBox.Show("Done!\n" + (watch.ElapsedMilliseconds / 1000) + " sec");
            if (IDErrors1.Count != 0 || IDErrors2.Count != 0)
            {
                MessageBox.Show("Id errors: " + IDErrors1.Count + ", " + IDErrors2.Count);
                List<string> iderrors = new List<string>();
                for (int i = 1; i < IDErrors1.Count; i++)
                {
                    iderrors.Add(IDErrors1[i].Range.Address + ", " + IDErrors2[i].Range.Address);
                }
                string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\IDErrors.json";
                using (StreamWriter writer = File.CreateText(file))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    var jsonString = JsonConvert.SerializeObject(iderrors);
                    //serializer.Serialize(writer, jsonString);
                    writer.Write(jsonString);
                }
            }
        }
        public void UpdateAllReferencesDB()
        {
            var watch = Stopwatch.StartNew();
            Main.instance.StopAll();
            foreach (ReferenceObject item in references.Values)
            {
                item.UpdateReferencesDB();
            }
            Main.instance.ResumeAll();
            watch.Stop();
            MessageBox.Show("Done!\n" + (watch.ElapsedMilliseconds / 1000) + " sec");
            if (IDErrors1.Count != 0 || IDErrors2.Count != 0)
            {
                MessageBox.Show("Id errors: " + IDErrors1.Count + ", " + IDErrors2.Count);
                List<string> iderrors = new List<string>();
                for (int i = 1; i < IDErrors1.Count; i++)
                {
                    iderrors.Add(IDErrors1[i].Range.Address + ", " + IDErrors2[i].Range.Address);
                }
                string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\IDErrors.json";
                using (StreamWriter writer = File.CreateText(file))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    var jsonString = JsonConvert.SerializeObject(iderrors);
                    //serializer.Serialize(writer, iderrors);
                    writer.Write(jsonString);
                }
            }
        }
        public void UpdateAllLevels(bool message = false)
        {
            foreach (ReferenceObject item in references.Values)
            {
                item.UpdateLevels();
            }
            if (message) MessageBox.Show("Done!");
        }

        // public void ClearAllDB()
        // {
        //     MessageBox.Show("Это может занять некоторое время! \nДождитесь сообщения об окончании.");
        //     GlobalMethods.ToLog("Инициализация очистки базы данных...");
        //     var watch = Stopwatch.StartNew();
        //     Main.instance.StopAll();
        //     foreach (ReferenceObject item in references.Values)
        //     {
        //         item.ClearAll();
        //     }
        //     Main.instance.ResumeAll();
        //     watch.Stop();
        //     GlobalMethods.ToLog("База данных очищена за " + (watch.ElapsedMilliseconds / 1000) + " sec.");
        //     MessageBox.Show("Готово!\n"+ (watch.ElapsedMilliseconds / 1000) + " sec.");
        // }
        public void ClearAllDB(bool message = true, SplashScreen splashScreen = null)
        {
            if (message == true) MessageBox.Show("Это может занять некоторое время! \nДождитесь сообщения об окончании.");
            GlobalMethods.ToLog("Инициализация очистки базы данных...");
            var watch = Stopwatch.StartNew();
            Main.instance.StopAll();
            foreach (ReferenceObject item in references.Values)
            {
                item.ClearAll();
                if (splashScreen != null)
                    splashScreen.UpdateText("Очистка " + item._name);
            }
            Main.instance.ResumeAll();
            watch.Stop();
            GlobalMethods.ToLog("База данных очищена за " + (watch.ElapsedMilliseconds / 1000) + " sec.");
            if (message == true) MessageBox.Show("Готово!\n"+ (watch.ElapsedMilliseconds / 1000) + " sec.");
        }

        public void UpdateMeterAllDB()
        {
            Main.instance.StopAll();
            foreach (ReferenceObject item in references.Values)
            {
                item.UpdateMeterAll();
            }
            Main.instance.ResumeAll();
        }

        public void CheckAllRanges()
        {
            Errors = new List<string>();

            var watch = Stopwatch.StartNew();

            Main.instance.StopAll();
            foreach (ReferenceObject ro in references.Values)
            {
                ro.Check();
            }
            Main.instance.ResumeAll();
            MessageBox.Show("Done!\n" + (watch.ElapsedMilliseconds / 1000) + " sec");

            string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\errors.json";
            using (StreamWriter writer = File.CreateText(file))
            {
                JsonSerializer serializer = new JsonSerializer();
                var jsonString = JsonConvert.SerializeObject(Errors);
                //serializer.Serialize(writer, Errors);
                writer.Write(jsonString);
            }
            Errors.Clear();
            Errors = null;
        }

        public void CheckDublicates()
        {
            Errors = new List<string>();

            var watch = Stopwatch.StartNew();
            Main.instance.StopAll();
            foreach (ReferenceObject ro in references.Values)
            {
                ReferenceObject co = references.Values.AsParallel().FirstOrDefault(n => n != ro && n.HasRange(ro.PS.Range)); //Main.instance.references.references.Values.Where(c => c != ro && c.HasRangePS(ro.PS.Range) == true).FirstOrDefault();

                if (co != null)
                {
                    Errors.Add("{" + co._name + "} and {" + ro._name + "}");
                }
            }
            Main.instance.ResumeAll();
            MessageBox.Show("Done!\n" + (watch.ElapsedMilliseconds / 1000) + " sec");

            string file = System.IO.Path.GetDirectoryName(Main.dir) + @"\DB\errors.json";
            using (StreamWriter writer = File.CreateText(file))
            {
                JsonSerializer serializer = new JsonSerializer();
                var jsonString = JsonConvert.SerializeObject(Errors);
                //serializer.Serialize(writer, Errors);
                writer.Write(jsonString);
            }
            Errors.Clear();
            Errors = null;
        }

        public void ReleaseAllComObjects()
        {
            references.Values.AsParallel().ForAll(r => r.ReleaseAllComObjects());
        }
    }
}

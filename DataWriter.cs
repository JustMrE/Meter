using System.Collections.Concurrent;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;

namespace Meter
{
    public static class DataWriter
    {
        public static void ProcessBatch(List<string> messages)
        {
            try
            {
                GlobalMethods.ToLog("Received " + messages.Count + " messages.");
                ConcurrentBag<PipeValue> result = DeserializeMessages<PipeValue>(messages);
                Write(result);
            }
            catch
            {
                GlobalMethods.ToLog("Ошибка записи для полученных данных.");
            }
        }

        static ConcurrentBag<T> DeserializeMessages<T>(List<string> messages)
        {
            ConcurrentBag<T> result = new ();
            List<Task> tasks = new List<Task>();

            foreach (string json in messages)
            {
                if (json != null)
                {
                    tasks.Add(Task.Run(() =>
                    {
                        try
                        {
                            T obj = JsonConvert.DeserializeObject<T>(json);
                            result.Add(obj);
                        }
                        catch (Exception)
                        {
                            
                        }
                    }));
                }
            }

            Task.WaitAll(tasks.ToArray());

            return result;
        }
        static void Write(ConcurrentBag<PipeValue> result)
        {
            Main.instance.StopAll();
            foreach (PipeValue pv in result)
            {
                GlobalMethods.ToLog("write to meter: " + pv.subjectName);
                if (!string.IsNullOrEmpty(pv.subjectName) && !string.IsNullOrEmpty(pv.level1Name) && !string.IsNullOrEmpty(pv.level2Name) && pv.day != null && !string.IsNullOrEmpty(pv.value))
                {
                    ReferenceObject ro = null;
                    if (Main.instance.references.references.TryGetValue(pv.subjectName, out ro))
                    {
                        if (ro != null)
                        {
                            if (ro.DB.childs.ContainsKey(pv.level1Name) && ro.DB.childs[pv.level1Name].childs.ContainsKey(pv.level2Name))
                            {
                                ro.WriteToDB(pv.level1Name, pv.level2Name, (int)pv.day, pv.value.Replace(",", "."));
                            }
                        }
                        else
                        {
                            GlobalMethods.ToLog("Не найден субъект " + pv.subjectName);
                        }
                    }
                    else
                    {
                        GlobalMethods.ToLog("Не найден субъект " + pv.subjectName);
                    }
                }
                else if (pv != null)
                {
                    if (pv.cod != null && pv.day != null && pv.value != null)
                    {
                        ReferenceObject ro = Main.instance.references.references.Values.AsParallel().Where(n => n.codPlan == pv.cod).FirstOrDefault();
                        if (ro != null)
                        {
                            ro.WriteToDB("план", "утвержденный", (int)pv.day, pv.value.Replace(",","."));
                        }
                        else
                        {
                            GlobalMethods.ToLogError("Не найден субъект с кодом плана " + pv.cod);
                            GlobalMethods.Err("Не найден субъект с кодом плана " + pv.cod);
                        }
                    }
                }
                else
                {
                    GlobalMethods.ToLog("Не достаточно данных для записи");
                }
            }
            Main.instance.ResumeAll();
        }
    }
}
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;

namespace Meter
{
    public static class DataWriter
    {
        public static void Write(string msg)
        {
            try
            {
                PipeValue pv = JsonConvert.DeserializeObject<PipeValue>(msg);

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
                            GlobalMethods.ToLog("Не найден субъект с кодом плана " + pv.cod);
                        }
                    }
                }
                else
                    {
                        GlobalMethods.ToLog("Не достаточно данных для записи");
                    }

            }
            catch
            {
                GlobalMethods.ToLog("Ошибка записи для полученных данных " + msg);
            }
        }
    }
}
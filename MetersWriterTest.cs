using System.Diagnostics;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;

namespace Meter
{
    public static class MetersWriterTest
    {
        public static void WriteMeters1(DateTime date)
        {
            string dict = Main.dir + @"\current\Словарь ТИ факт.xlsx";
            string path = "H1";
            // string path = "H2";
            string fName = "I1";
            string fileWSName = "сч";
            int day = date.Day;
            Excel.Application xlApp;
            Excel.Workbook w1;
            Excel.Worksheet ws1;
            Excel.Range r;
            string wa1;
            Dictionary<string, Dictionary<string,string>> dict2 = new();
            List<Subject> subjects = new List<Subject>();

            var watch = Stopwatch.StartNew();
            GlobalMethods.ToLog("Считывание словаря: " + dict);
            Forms.SplashScreen splashScreen = new();
            splashScreen.Show();

            splashScreen.UpdateLabel("Запись счетчиков: ");
            splashScreen.UpdateText("cчитывание словаря");
            xlApp = new Excel.ApplicationClass
            {
                Visible = false
            };
            w1 = xlApp.Workbooks.Open(dict);
            for (int i = 3; i <= w1.Worksheets.Count; i++)
            {
                dict2.Add((string)((Excel.Worksheet)w1.Worksheets[i]).Range[path].Value + (string)((Excel.Worksheet)w1.Worksheets[i]).Range[fName].Value, ReadTIDict((Excel.Worksheet)w1.Worksheets[i]));
            }
            w1.Close(false);
            w1 = null;
            GlobalMethods.ToLog("Считывание словаря завершено: " + dict);
            foreach (string file in dict2.Keys)
            {
                try 
                {
                    w1 = xlApp.Workbooks.Open(file, false);
                    ws1 = (Excel.Worksheet)w1.Worksheets[fileWSName];
                    GlobalMethods.ToLog("Считывание данных: " + file);
                    splashScreen.UpdateText("Считывание данных: " + file.Substring(file.LastIndexOf("\\") + 1));

                    foreach (string item in dict2[file].Keys)
                    {
                        string[] k = item.Split("/");
                        string[] v = dict2[file][item].Split("/");

                        GlobalMethods.ToLog("Считывание данных: " + item);
                        try
                        {
                            r = ws1.Range["A:A"].Find(k[0], LookAt: Excel.XlLookAt.xlWhole).MergeArea;
                            int c1, c2, r1, r2;
                            c1 = ((Excel.Range)r.Cells[1, 1]).Column;
                            r1 = ((Excel.Range)r.Cells[1, 1]).Row;
                            c2 = ((Excel.Range)r.Cells[r.Cells.Rows.Count, r.Cells.Columns.Count]).Column;
                            r2 = ((Excel.Range)r.Cells[r.Cells.Rows.Count, r.Cells.Columns.Count]).Row;
                            c1 += 1;
                            c2 += 1;
                            
                            string adr = ws1.Range[ws1.Cells[r1, c1], ws1.Cells[r2, c2]].Address;
                            if (ws1.Range[adr].Find(k[1], After: ws1.Range[adr].Cells[ws1.Range[adr].Cells.Count], LookAt: Excel.XlLookAt.xlPart) != null)
                            {
                                adr = ws1.Range[adr].Find(k[1], After: ws1.Range[adr].Cells[ws1.Range[adr].Cells.Count], LookAt: Excel.XlLookAt.xlPart).Address;
                                r1 = ws1.Range[adr].Row;
                                c1 = 3 + day;
                                adr = ((Excel.Range)ws1.Cells[r1, c1]).Address;
                                
                                if (ws1.Range[adr].Value != null)
                                {
                                    subjects.Add(new Subject()
                                    {
                                        subjectName = v[0],
                                        level1Name = v[1],
                                        level2Name = "счетчик",
                                        val = ((double)ws1.Range[adr].Value).ToString()
                                    });
                                }
                            }
                        }
                        catch (System.Exception e)
                        {
                            GlobalMethods.Err(e + " " + "{" + k[0] + "}");
                            GlobalMethods.ToLog(e + " " + "{" + k[0] + "}");
                        }
                        GlobalMethods.ToLog("Считывание данных завершено: " + item);
                    }
                    w1.Close(false);
                    w1 = null;
                    ws1 = null;
                    r = null;
                }
                catch (Exception e) 
                {
                    GlobalMethods.ToLog("Файл: " + file + " ошибка: " + e.Message);
                    splashScreen.UpdateText(e.Message);
                    Thread.Sleep(1000);
                }
            }

            ws1 = null;
            w1 = null;
            r = null;
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;

            splashScreen.UpdateText("запись данных");
            Main.instance.StopAll();
            ReferenceObject ro = null;
            foreach (Subject s in subjects)
            {
                splashScreen.UpdateText("запись данных " + s.subjectName);
                if (Main.instance.references.references.TryGetValue(s.subjectName, out ro))
                {
                    if (ro != null)
                    {
                        if (ro.DB.childs.ContainsKey(s.level1Name) && ro.DB.childs[s.level1Name].childs.ContainsKey(s.level2Name))
                        {
                            ro.WriteToDB(s.level1Name, s.level2Name, (int)day, s.val.Replace(",", "."));
                        }
                    }
                    else
                    {
                        GlobalMethods.ToLog("Не найден субъект " + s.subjectName);
                    }
                }
                else
                {
                    GlobalMethods.ToLog("Не найден субъект " + s.subjectName);
                }
            }
            Main.instance.ResumeAll();

            splashScreen.Close();
            watch.Stop();
            MessageBox.Show("Done!\n" + (watch.ElapsedMilliseconds / 1000) + " sec");
        }

        private static Dictionary<string, string> ReadTIDict(Excel.Worksheet ws1)
        {
            Dictionary<string, string> tiDict = new ();
            Excel.Range r = ws1.Range["A2"];
            while (r.Value != null)
            {
                if (r.Value != null && r.Offset[0, 1].Value != null && r.Offset[0, 2].Value != null && r.Offset[0, 3].Value != null)
                {
                    string k = r.Value + "/" + r.Offset[0, 1].Value;
                    string v = r.Offset[0, 2].Value + "/" + r.Offset[0, 3].Value;
                    tiDict.Add(k, v);
                }
                r = r.Offset[1, 0];
            }
            return tiDict;
        }
    }
    struct Subject
    {
        public string subjectName, level1Name, level2Name, val;
    }
}
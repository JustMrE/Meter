using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;
using Meter.Forms;

namespace Meter
{
    public static class MetersWriterTest
    {
        public static void WriteOldMeters(DateTime date)
        {
            Excel.Application xlApp;
            Excel.Workbook w1;
            Excel.Worksheet ws1;
            List<string> mArray = new List<string>(){"январь_н", "февраль_н", "март_н", "апрель_н", "май_н", "июнь_н", "июль_н", "август_н", "сентябрь_н", "октябрь_н", "ноябрь_н", "декабрь_н"};
            int d, day;
            string wsName = mArray[date.Month - 1];
            day = date.Day;
            string val;
            d = 0;
            if (date.Day <= 10)
            {
                d = date.Day;
            }
            else if (date.Day <= 20)
            {
                d = date.Day + 1;
            }
            else if (date.Day <= 31)
            {
                d = date.Day + 3;
            }
            List<OldMeter> oldMeter = CreateValuesOldMeter();

            Forms.SplashScreen splashScreen = new();
            splashScreen.Show();

            splashScreen.UpdateLabel("Запись счетчиков: ");
            splashScreen.UpdateText("Запись счетчиков");

            xlApp = new Excel.ApplicationClass
            {
                Visible = false
            };
            // w1 = xlApp.Workbooks.Open(@"X:\OPER\счетчики.xls");
            w1 = xlApp.Workbooks.Open(MeterSettings.Instance.OldMeterFile);
            ws1 = w1.Worksheets[wsName] as Excel.Worksheet;

            Main.instance.StopAll();
            ReferenceObject ro = null;
            foreach (OldMeter s in oldMeter)
            {
                splashScreen.UpdateText("запись данных " + s.subjectName);
                if (Main.instance.references.references.TryGetValue(s.subjectName, out ro))
                {
                    if (ro != null)
                    {
                        if (ro.DB.childs.ContainsKey(s.level1Name) && ro.DB.childs[s.level1Name].childs.ContainsKey(s.level2Name))
                        {
                            if (s.subjectName.Contains("Мын Арал тяга"))
                            {
                                val = (((double)ws1.Range[s.adr].Offset[d, 0].Value)/1000).ToString();
                            }
                            else
                            {
                                try
                                {
                                    val = ((double)ws1.Range[s.adr].Offset[d, 0].Value).ToString();
                                }
                                catch (System.Exception e)
                                {
                                    val = "";
                                    GlobalMethods.Err("{" + s.level1Name + "} {" + s.level2Name + "} " + e.Message);
                                }
                            }
                            ro.WriteToDB(s.level1Name, s.level2Name, (int)day, val.Replace(",", "."));
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
            w1.Close(false);
            splashScreen.Close();
        }

        static List<OldMeter> CreateValuesOldMeter()
        {
            return new()
            {
                new ()
                {
                    adr = "FB178",
                    subjectName = "Л5143 Фрунзе",
                    level1Name = "прием",
                    level2Name = "счетчик"
                },
                new ()
                {
                    adr = "FC178",
                    subjectName = "Л5143 Фрунзе",
                    level1Name = "отдача",
                    level2Name = "счетчик"
                },
                new ()
                {
                    adr = "AD178",
                    subjectName = "ТП1 Жидели тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AE178",
                    subjectName = "ТП2 Жидели тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AF178",
                    subjectName = "ТП3 Жидели тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AG178",
                    subjectName = "ТП1 Кияхты тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AH178",
                    subjectName = "ТП2 Кияхты тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AI178",
                    subjectName = "ТП3 Кияхты тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AO50",
                    subjectName = "ТП1 Чиганак тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AP50",
                    subjectName = "ТП2 Чиганак тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AQ50",
                    subjectName = "ТП3 Чиганак тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AR50",
                    subjectName = "ТП1 Мын Арал тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AS50",
                    subjectName = "ТП2 Мын Арал тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AT50",
                    subjectName = "ТП3 Мын Арал тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AU50",
                    subjectName = "ТП1 Узун Агач тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AV50",
                    subjectName = "ТП2 Узун Агач тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AW50",
                    subjectName = "ТП3 Узун Агач тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AX50",
                    subjectName = "ТП1 Медеу тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AY50",
                    subjectName = "ТП2 Медеу тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "BK310",
                    subjectName = "с.н. Мойнак ГЭС МГЭС",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "EA589",
                    subjectName = "Т1 Eneverse Kunkuat СЭС Капшагай 100",
                    level1Name = "прием",
                    level2Name = "счетчик"
                },
                new ()
                {
                    adr = "EB589",
                    subjectName = "Т1 Eneverse Kunkuat СЭС Капшагай 100",
                    level1Name = "отдача",
                    level2Name = "счетчик"
                },
                new ()
                {
                    adr = "EC589",
                    subjectName = "Т2 Eneverse Kunkuat СЭС Капшагай 100",
                    level1Name = "прием",
                    level2Name = "счетчик"
                },
                new ()
                {
                    adr = "ED589",
                    subjectName = "Т2 Eneverse Kunkuat СЭС Капшагай 100",
                    level1Name = "отдача",
                    level2Name = "счетчик"
                },
                new ()
                {
                    adr = "BC136",
                    subjectName = "Генерация Каскад Каратальских ГЭС ККГЭС",
                    level1Name = "отдача",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "BM136",
                    subjectName = "Генерация Каскад Каратальских ГЭС ККГЭС",
                    level1Name = "отдача",
                    level2Name = "корректировка факт"
                },
                new ()
                {
                    adr = "BO136",
                    subjectName = "с.н. Каскад Каратальских ГЭС ККГЭС",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "EJ136",
                    subjectName = "с.н. Каскад Каратальских ГЭС ККГЭС",
                    level1Name = "прием",
                    level2Name = "корректировка факт"
                },
                new ()
                {
                    adr = "EK136",
                    subjectName = "с.н. Аксу ГЭС",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "K178",
                    subjectName = "Л133А Шу 110 Тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "L178",
                    subjectName = "Л134А Шу 110 Тяга",
                    level1Name = "прием",
                    level2Name = "ручное"
                },
                new ()
                {
                    adr = "AF94",
                    subjectName = "Л2313 Сары Шаган тяга",
                    level1Name = "прием",
                    level2Name = "счетчик"
                },
                new ()
                {
                    adr = "AE94",
                    subjectName = "Л2313 Сары Шаган тяга",
                    level1Name = "отдача",
                    level2Name = "счетчик"
                },
            };
        }

        struct OldMeter
        {
            public string adr, subjectName, level1Name, level2Name;
        }
        public static void WriteMeters1(DateTime date)
        {
            string dict = MeterSettings.Instance.DBDir + @"\current\Словарь ТИ факт.xlsx";
            string path;
            string archPath = "";
            path = MeterSettings.Instance.OtherMetersPath;
            // if (false)
            // {
            //     path = "H2";
            //     archPath = date.Year + "\\" + date.ToString("MMMM", GlobalMethods.culture) + "\\";
            // }
            // else
            // {
            //     path = "H1";
            // }
            
            // path = "H1";
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
                dict2.Add(path + "\\" + (string)((Excel.Worksheet)w1.Worksheets[i]).Range[fName].Value, ReadTIDict((Excel.Worksheet)w1.Worksheets[i]));
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
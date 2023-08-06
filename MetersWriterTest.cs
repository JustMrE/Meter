using System.Diagnostics;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;

namespace Meter
{
    public static class MetersWriterTest
    {
        public static void WriteMeters(DateTime date)
        {
            Main.instance.StopAll();
            string file = "";
            string dict = Main.dir + @"\current\Словарь ТИ факт.xlsx";
            // string dictWSName = "Записать ОРУ МГЭС";
            string fileWSName = "сч";
            int day = date.Day;
            Excel.Workbook w1;
            Excel.Worksheet ws1;

            var watch = Stopwatch.StartNew();
            
            Dictionary<string,string> tiDict = new Dictionary<string, string>();
            List<Subject> subjects = new List<Subject>();
            List<string> wsNames = new List<string>();
            w1 = Main.instance.xlApp.Workbooks.Open(dict);

            for (int i = 3; i < w1.Worksheets.Count; i++)
            {
                wsNames.Add(((Excel.Worksheet)w1.Worksheets[i]).Name);
            }
            w1.Close(false);
            w1 = null;

            foreach (string dictWSName in wsNames)
            {
                tiDict.Clear();
                subjects.Clear();
                w1 = Main.instance.xlApp.Workbooks.Open(dict);
                ws1 = (Excel.Worksheet)w1.Worksheets[dictWSName];
                file = (string)ws1.Range["H1"].Value + (string)ws1.Range["I1"].Value;
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
                
                w1.Close(false);
                w1 = null;
                ws1 = null;
                r = null;

                w1 = Main.instance.xlApp.Workbooks.Open(file, false);
                ws1 = (Excel.Worksheet)w1.Worksheets[fileWSName];

                foreach (string item in tiDict.Keys)
                {
                    string[] k = item.Split("/");
                    string[] v = tiDict[item].Split("/");

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
                w1.Close(false);
                w1 = null;
                ws1 = null;
                r = null;

                ReferenceObject ro = null;
                foreach (Subject s in subjects)
                {
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
            }
            watch.Stop();
            MessageBox.Show("Done!\n" + (watch.ElapsedMilliseconds / 1000) + " sec");
        }
    }

    struct Subject
    {
        public string subjectName, level1Name, level2Name, val;
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Emcos;
using System.Net;
using System.Data;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Globalization;
using System.Collections.Concurrent;
using static Meter.CustomCellMenu;

namespace Meter.Forms
{
    partial class NewMenuBase
    {
        [DllImport("user32.dll", SetLastError = true)]
        protected static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        protected static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOOLWINDOW = 0x80;
        public IntPtr formHwnd;

        private static object[,] oldValsArray;
        private static string oldVal;
        protected static string token = "";
        //protected Color activeColor;
        public Color activeColor;
        public static Excel.Range _activeRange;
        public static bool restartFlag = false;
        public static HashSet<string> editedFormulas = new ();
        public static HashSet<string> selectedButtons = new ();

        
        private void RegexSearch()
        {
            RegexOptions ro = checkBox1.Checked ? RegexOptions.None : RegexOptions.IgnoreCase;
            if (this.tbSearch.Text != "")
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    try
                    {
                        listBox1.Items.Clear();
                        string search = this.tbSearch.Text;
                        search = search.Replace("*", @".*");

                        var deviceIds = Main.instance.references.references.Values.AsEnumerable();
                        var matchingIds = deviceIds.Where(id => Regex.IsMatch(id._name, pattern: search, ro)).Select(n => new ListViewItem(n._name){ToolTipText = n._name}).ToArray();
                        listBox1.Items.AddRange(matchingIds.ToArray());
                    }
                    catch (ArgumentException)
                    {
                        this.Invoke((MethodInvoker)(() =>
                        {
                            listBox1.Items.Clear();
                            listBox1.Items.AddRange(Main.instance.references.references.Keys.OrderBy(m => m).Select(str => new ListViewItem(str){ToolTipText = str}).ToArray());
                        }));
                    }

                }));
            }
            else
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    listBox1.Items.Clear();
                    listBox1.Items.AddRange(Main.instance.references.references.Keys.OrderBy(m => m).Select(str => new ListViewItem(str){ToolTipText = str}).ToArray());
                }));
            }

        }
        private void RegexSearchHeads()
        {
            RegexOptions ro = checkBox1.Checked ? RegexOptions.None : RegexOptions.IgnoreCase;
            if (this.tbSearch.Text != "")
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    try
                    {
                        listBox1.Items.Clear();
                        string search = this.tbSearch.Text;
                        search = search.Replace("*", @".*");

                        var deviceIds = HeadReferences.idDictionary.Values.ToList();
                        var matchingIds = deviceIds.Where(id => Regex.IsMatch(id._name, pattern: search, ro)).Select(n => new ListViewItem(n._name){ToolTipText = n._name}).ToArray();
                        listBox1.Items.AddRange(matchingIds.ToArray());
                    }
                    catch (ArgumentException)
                    {
                        this.Invoke((MethodInvoker)(() =>
                        {
                            listBox1.Items.Clear();
                            listBox1.Items.AddRange(HeadReferences.idDictionary.Values.Select(n => n._name).OrderBy(m => m).Select(n => new ListViewItem(n){ToolTipText = n}).ToArray());
                        }));
                    }

                }));
            }
            else
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    listBox1.Items.Clear();
                    listBox1.Items.AddRange(HeadReferences.idDictionary.Values.Select(n => n._name).OrderBy(m => m).Select(n => new ListViewItem(n){ToolTipText = n}).ToArray());
                }));
            }

        }
        public void ChangeVisible(bool visible)
        {
            Action<bool> action = (bool vis) => { this.Visible = vis; };
            if (InvokeRequired)
            {
                Invoke(action, visible);
            }
            else
            {
                action(visible);
            }
        }
        public virtual void SetRects(int left, int top, int width, int height)
        {
            Action<int, int, int, int> action = SetPosition;
            if (InvokeRequired)
            {
                Invoke(action, left, top, width, height);
            }
            else
            {
                action(left, top, width, height);
            }
        }
        private void SetPosition(int left, int top, int width, int height)
        {
            Location = new System.Drawing.Point(left, top);
            Width = width;
            Height = height;
        }
        public virtual void FormClose()
        {
            Action action = Close;
            if (InvokeRequired)
            {
                Invoke(action);
            }
            else
            {
                action();
            }
        }
    
        public virtual void PreContextMenu()
        {
            // cb = Main.instance.xlApp.CommandBars[Main.menuIndexes[0]];
            selectedButtons.Clear();
            // RecreateCustomContextMenu();
            ContextMenu();
        }
        public virtual void ContextMenu()
        {
            GlobalMethods.ToLog("Открыто контекстное меню");
            selectedButtons.Add("Копировать");
            IDataObject clipboardData = null;
            Invoke(new Action(() => clipboardData = Clipboard.GetDataObject()));
            if (clipboardData != null && clipboardData.GetFormats().Length > 0) selectedButtons.Add("Вставить");
            if (Main.instance.colors.subColors.ContainsValue(activeColor))
            {
                HeadObject ho = Main.instance.heads.HeadByRange(_activeRange);
                if (ho != null)
                {
                    if (ho.GetParent != null && ho.LastColumn.Column != ho.GetParent.LastColumn.Column)
                    {
                        if (ho.HasIndent(IndentDirection.right) == true)
                        {
                            selectedButtons.Add("Удалить отступ");
                        }
                        else
                        {
                            selectedButtons.Add("Добавить отступ");
                        }
                    }
                }
                
            }
            if (Main.instance.colors.mainTitle.ContainsValue(activeColor))
            {
                selectedButtons.Add("Изменить main");
                selectedButtons.Add("Сбросить");
            }
            if (Main.instance.colors.mainSubtitle.ContainsValue(activeColor))
            {
                selectedButtons.Add("Изменить mainSubtitle");
                selectedButtons.Add("Сбросить mainSubtitle");
                selectedButtons.Add("Ввести корректировку");
                if (RangeReferences.activeTable.DB.HasItem("по счетчику"))
                {
                    selectedButtons.Add("Ввести показания счетчика");
                }
                if (RangeReferences.activeTable.DB.HasItem("аскуэ"))
                {
                    selectedButtons.Add("Записать данные из EMCOS");
                    selectedButtons.Add("Очистить данные из EMCOS");
                }
                
            }
            if (Main.instance.colors.extraTitle.ContainsValue(activeColor))
            {
                selectedButtons.Add("Изменить extra");
                selectedButtons.Add("Скрыть");
            }
            if (Main.instance.colors.main.ContainsValue(activeColor) && Main.instance.colors.main["subject"] != activeColor)
            {
                selectedButtons.Add("Показать");
                if (RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].HasItem("формула"))
                {
                    selectedButtons.Add("Изменить формулу");
                }
                // if (!RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].HasItem("счетчик"))
                // {
                //     selectedButtons.Add("Добавить по показаниям счетчика");
                // }
                // else
                // {
                //     selectedButtons.Add("Удалить по показаниям счетчика");
                // }
                selectedButtons.Add("Зависимые формулы");
            }
            if (activeColor == Main.instance.colors.main["subject"])
            {
                selectedButtons.Add("Выделить");
                if (RangeReferences.activeTable.DB.HasItem("аскуэ"))
                {
                    selectedButtons.Add("Записать данные из EMCOS (Все дни)");
                    selectedButtons.Add("Очистить данные из EMCOS (Все дни)");
                }
                if (RangeReferences.activeTable.DB.HasItem("по счетчику"))
                {
                    selectedButtons.Add("Изменить коэффициент счетчика");
                }
            }
        }
        public virtual void DeactivateSheet()
        {
            ChangeVisible(false);
        }
        public virtual void ActivateSheet(object sh)
        {
            ChangeVisible(true);
        }
        public virtual void RightClick(Excel.Range range)
        {
            GlobalMethods.ToLog("Нажата правая кнопка мыши на ячейке " + range.Address);
            _activeRange = range;
            activeColor = _activeRange.ColorRGB();
            Main.instance.references.ActivateTable(range);
            PreContextMenu();
            MyContextMenu();
            cb.ShowPopup();
        }
        public virtual void DblClick(Excel.Range range)
        {
            SelectSubject(range);
        }
        public virtual void SlectionChanged(Excel.Range range)
        {
            _activeRange = range;
            Main.instance.references.ActivateTable(range);
            if (range.Formula is string)
            {
                oldValsArray = null;
                oldVal = (string)range.Formula;
            }
            else if (range.Formula is object[,])
            {
                oldVal = null;
                oldValsArray = (object[,])range.Formula;
            }
            
            if (Main.instance.zoom != (double)Main.instance.xlApp.ActiveWindow.Zoom)
            {
                Main.instance.zoom = (double)Main.instance.xlApp.ActiveWindow.Zoom;
                GlobalMethods.CalculateFormsPositions();
            }
            restartFlag = false;
        }
        public virtual void CellValueChanged(Excel.Range range)
        {
            if (range.Address[false,false] == "A1")
            {
                return;
            }
            Main.instance.xlApp.EnableEvents = false;
            Main.instance.xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

            var coloredCellsToProcess = range.Cast<Excel.Range>().Where(cell =>
            {
                Color c = ColorsData.GetRangeColor(cell);
                return c == Main.instance.colors.mainSubtitle["ручное"] || 
                c == Main.instance.colors.mainSubtitle["корректировка"] || 
                c == Main.instance.colors.mainSubtitle["корректировка факт"] ||
                c == Main.instance.colors.mainSubtitle["счетчик"] ||
                c == Main.instance.colors.extraSubtitle["РУЧНОЕ"] ||
                c == Main.instance.colors.extraSubtitle["КОРРЕКТИРОВКА"] ||
                c == Main.instance.colors.extraSubtitle["КОРРЕКТИРОВКА ФАКТ"] ||
                c == Main.instance.colors.extraSubtitle["СЧЕТЧИК"]; 
            });

            foreach (Excel.Range cell in coloredCellsToProcess)
            {   
                ReferenceObject ro = Main.instance.references[cell];
                if (ro != null)
                {
                    string nameL1 = ro.GetL1(cell)._name;
                    string nameL2 = ro.GetL2(cell, nameL1)._name;
                    // string nameL2 = Main.instance.colors.NameByColor(ColorsData.GetRangeColor(cell));
                    int? day = ro.PS.childs[nameL1].childs[nameL2].DayByRange(cell);
                    if (day != null)
                    {
                        string dbNameL2 = Main.instance.colors.NameByColor(ColorsData.GetRangeColor(cell));
                        string val = cell.Formula != null ? cell.Formula.ToString() : "";
                        ro.WriteToDB(nameL1, dbNameL2, (int)day, val);
                    }
                }
                
            }

            DontEdit(range);
            #region Old
                // {
                //     if (range.Cells.Count != _activeRange.Cells.Count)
                //     {
                //         int r, c;
                //         Excel.Range r1 = (Excel.Range)_activeRange.Cells[1, 1];
    
                //         r = 1 + (range.Row - r1.Row);
                //         c = 1 + (range.Column - r1.Column);
                //         oldVal = (string)oldValsArray[r, c];
                //         Marshal.ReleaseComObject(r1);
                //     }
                //     if (range.Cells.Count > 1) 
                //     {
                //         // oldValsArray = (object[,])range.Formula;
                //         // foreach (Excel.Range cell in range.Cells)
                //         // {
                //         //     Color c = ColorsData.GetRangeColor(cell);
                //         //     if (c == Main.instance.colors.mainSubtitle["ручное"] || 
                //         //         c == Main.instance.colors.mainSubtitle["корректировка"] || 
                //         //         c == Main.instance.colors.mainSubtitle["корректировка факт"] ||
                //         //         c == Main.instance.colors.mainSubtitle["счетчик"] ||
                //         //         c == Main.instance.colors.extraSubtitle["РУЧНОЕ"] ||
                //         //         c == Main.instance.colors.extraSubtitle["КОРРЕКТИРОВКА"] ||
                //         //         c == Main.instance.colors.extraSubtitle["КОРРЕКТИРОВКА ФАКТ"] ||
                //         //         c == Main.instance.colors.extraSubtitle["СЧЕТЧИК"])
                //         //     {
                //         //         _activeRange = cell;
                //         //         Main.instance.references.ActivateTable(_activeRange);
                //         //         if (RangeReferences.activeTable != null)
                //         //         {
                //         //             int? day = RangeReferences.activeTable.ActiveDay();
                //         //             if (day != null)
                //         //             {
                //         //                 WriteToDB(_activeRange, (int)day, false);
                //         //             }
                //         //         }
    
                //         //     }
                //         // }
                //         // _activeRange = range;
                //         DontEditRange(range);
                //     }
                //     else
                //     {
                //         Color c = ColorsData.GetRangeColor(range);
                //         if (c == Main.instance.colors.mainSubtitle["ручное"] || 
                //             c == Main.instance.colors.mainSubtitle["корректировка"] || 
                //             c == Main.instance.colors.mainSubtitle["корректировка факт"] ||
                //             c == Main.instance.colors.mainSubtitle["счетчик"] ||
                //             c == Main.instance.colors.extraSubtitle["РУЧНОЕ"] ||
                //             c == Main.instance.colors.extraSubtitle["КОРРЕКТИРОВКА"] ||
                //             c == Main.instance.colors.extraSubtitle["КОРРЕКТИРОВКА ФАКТ"] ||
                //             c == Main.instance.colors.extraSubtitle["СЧЕТЧИК"])
                //         {
                //             int? day = RangeReferences.activeTable.ActiveDay();
                //             if (day != null)
                //             {
                //                 string? name = Main.instance.colors.NameByColor(c);
                //                 WriteToDB(_activeRange, (int)day, true, L2: name);
                //             }
                //         }
                //         else
                //         {
                //             DontEdit(range);
                //         }
                //     }
                // }
            #endregion
            Main.instance.xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Main.instance.xlApp.EnableEvents = true;
        }

        private void DontEdit(Excel.Range rng, int? day = null)
        {
            if (oldVal != null)
            {
                rng.Formula = oldVal;
            }
            else if (oldValsArray != null)
            {
                int numRows = oldValsArray.GetLength(0);
                int numCols = oldValsArray.GetLength(1);

                if (numRows == rng.Rows.Count && numCols == rng.Columns.Count)
                {
                    rng.Formula = oldValsArray;
                }
            }
        }
        private void NewMenuBase_FormClosed(object sender, FormClosedEventArgs e)
        {
            Marshal.ReleaseComObject(CustomCellMenu.cb);
            if (_activeRange != null) Marshal.ReleaseComObject(_activeRange);
        }
        
        public virtual void GotoDB()
        {
            //MessageBox.Show(RangeReferences.activeTable._name);
            Main.instance.wsDb.Activate();
            RangeReferences.activeTable.DB.Range.Select();
        }
        public virtual void SelectSubject()
        {
            if (RangeReferences._activeObject != null) RangeReferences._activeObject.Range.Select();
        }
        protected virtual void SelectSubject(Excel.Range range)
        {
            Main.instance.references.ActivateTable(range);
            if (RangeReferences._activeObject != null) RangeReferences._activeObject.Range.Select();
        }
        protected virtual void SelectSubject(string name)
        {
            if (Main.instance.references.references.ContainsKey(name)) Main.instance.references.references[name].PS.Range.Select();
        }
        
        public void HideType()
        {
            RangeReferences.activeTable._activeChild._activeChild._activeChild.Remove();
        }
        
        public void EmcosSelect()
        {
            using (EmcosPicker emcosPicker = new EmcosPicker())
            {
                var result = emcosPicker.ShowDialog();
                if (result == DialogResult.OK)
                {
                    RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].emcosID = emcosPicker.id;
                    RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].emcosMLID = emcosPicker.ML_ID;
                }
            }
        }
        public void EmcosRemove()
        {
            RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].emcosID = null;
            RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].emcosMLID = null;
        }
        public void EmcosWrite(DateTime from, DateTime to, ReferenceObject ro = null)
        {
            if (EmcosMethods.LoginToEmcos())
            {
                SplashScreen splashScreen = new();
                splashScreen.Show();
                splashScreen.UpdateLabel("Запись данных из АСКУЭ");
                ReferenceObject[] ranges;
                if (ro == null)
                {
                    ranges = Main.instance.references.references.Values.Where(n => n.HasEmcosID == true).ToArray();
                }
                else
                {
                    ranges = new [] {ro};
                }
                ConcurrentBag<ConcurrentBag<EMCOSObject>> emcosValues = new ();
                List<DateTime> dates = new List<DateTime>();

                for (DateTime d = from; d.Day <= to.Day; d = d.AddDays(1))
                {
                    dates.Add(d);
                }
                
                splashScreen.UpdateText("считывание данных из АСКУЭ");
                Main.instance.StopAll();
                Parallel.ForEach(ranges, r => 
                {
                    ConcurrentBag<EMCOSObject> es = new ();
                    Parallel.ForEach(r.DB.childs.Values, v => 
                    {
                        if (v.HasItem("аскуэ") && v.emcosID != null)
                        {
                            EMCOSObject eo = new ();
                            eo.name = r._name;
                            eo.dbid = v.childs["аскуэ"].ID;
                            eo.psid = r.PS.childs[v._name].childs["аскуэ"].ID;
                            eo.values = new();
                            eo.flags = new();
                            Parallel.ForEach(dates, d =>
                            {
                                float? floatVal = null;
                                string val;
                                bool flag = EmcosMethods.GetValue(d, d, v, ref floatVal);
                                if (floatVal != null)
                                {
                                    floatVal = floatVal / 1000f;
                                    val = floatVal.ToString().Replace(",", ".");
                                }
                                else
                                {
                                    val = "0";
                                }
                                eo.values.TryAdd(d, val);
                                eo.flags.TryAdd(d, flag);
                            });
                            es.Add(eo);
                        }
                    });
                    emcosValues.Add(es);
                });

                foreach (var s in emcosValues)
                {
                    foreach (var t in s)
                    {
                        splashScreen.UpdateText("запись данных из АСКУЭ: " + t.name);
                        foreach (var d in t.values.Keys)
                        {
                            splashScreen.UpdateLabel("Запись данных из АСКУЭ за " + d.ToString("dd.MM.yy") + ": ");
                            ((ChildObject)RangeReferences.idDictionary[t.dbid]).WriteValue(d.Day, t.values[d]);
                            if (t.flags[d] == true)
                            {
                                ((ChildObject)RangeReferences.idDictionary[t.psid]).AddNote("Неполные данные", d.Day);
                            }
                        }
                    }
                }
                Main.instance.ResumeAll();
                splashScreen.Close();
                MessageBox.Show("Done!");
            }
            else
            {
                MessageBox.Show("Done!\nCan't login to EMCOS.");
            }
        }
        public void EmcosClear(DateTime from, DateTime to, ReferenceObject ro = null)
        {
            
            SplashScreen splashScreen = new();
            splashScreen.Show();
            splashScreen.UpdateLabel("Очистка данных из АСКУЭ");
            ReferenceObject[] ranges;
            if (ro == null)
            {
                ranges = Main.instance.references.references.Values.Where(n => n.HasEmcosID == true).ToArray();
            }
            else
            {
                ranges = new [] {ro};
            }
            ConcurrentBag<ConcurrentBag<EMCOSObject>> emcosValues = new ();
            List<DateTime> dates = new List<DateTime>();

            for (DateTime d = from; d.Day <= to.Day; d = d.AddDays(1))
            {
                dates.Add(d);
            }
            Main.instance.StopAll();
            Parallel.ForEach(ranges, r => 
            {
                ConcurrentBag<EMCOSObject> es = new ();
                Parallel.ForEach(r.DB.childs.Values, v => 
                {
                    if (v.HasItem("аскуэ") && v.emcosID != null)
                    {
                        EMCOSObject eo = new ();
                        eo.name = r._name;
                        eo.dbid = v.childs["аскуэ"].ID;
                        eo.psid = r.PS.childs[v._name].childs["аскуэ"].ID;
                        eo.values = new();
                        Parallel.ForEach(dates, d =>
                        {
                            eo.values.TryAdd(d, "");
                        });
                        es.Add(eo);
                    }
                });
                emcosValues.Add(es);
            });

            foreach (var s in emcosValues)
            {
                foreach (var t in s)
                {
                    splashScreen.UpdateText("очистка данных из АСКУЭ: " + t.name);
                    foreach (var d in t.values.Keys)
                    {
                        splashScreen.UpdateLabel("Очистка данных из АСКУЭ за " + d.ToString("dd.MM.yy") + ": ");
                        ((ChildObject)RangeReferences.idDictionary[t.dbid]).WriteValue(d.Day, t.values[d]);
                        ((ChildObject)RangeReferences.idDictionary[t.psid]).ClearNote(d.Day);
                        
                    }
                }
            }
            Main.instance.ResumeAll();
            splashScreen.Close();
            MessageBox.Show("Done!");
            
        }
        public void MyContextMenu()
        {
            
            foreach (CommandBarControl item in cb.Controls)
            {
                item.Visible = false;
            }

            foreach (string bn in selectedButtons)
            {
                cb.Controls[CommandBarIndexes[bn]].Visible = true;
                if (CommandBarActions.ContainsKey(bn)) CommandBarActions[bn].Invoke((CommandBarPopup)cb.Controls[CommandBarIndexes[bn]]);
            }
        }
        public void ClearContextMenu()
        {
            foreach (CommandBarControl item in CustomCellMenu.cb.Controls)
            {
                item.Visible = false;
            }
        }
        public void Test()
        {
            Invoke(new MethodInvoker(() =>
            {
                // Создаем и открываем форму в потоке2
                FormulaEditor form = new FormulaEditor(ref RangeReferences.activeTable, RangeReferences.ActiveL1);
                form.Show();
            }));
        }
    }
}

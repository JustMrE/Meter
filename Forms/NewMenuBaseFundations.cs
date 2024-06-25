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
            if (textBox1.Text == "dontsave")
            {
                MeterSettings.Instance.CloseAutoSave = true;
            }
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
                    selectedButtons.Add("Записать из EMCOS");
                    selectedButtons.Add("Очистить из EMCOS");
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
            ResetContextMenu();
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
        
        protected void SpecialMenuMain()
        {
            cb.FindControl(Tag: "Special").Visible = true;
            /*Action action;

            CommandBarPopup p = (CommandBarPopup)CustomCellMenu.cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Special";
            
            if (Main.instance.colors.main["subject"] == activeColor)
            {
                //AddButtonToPopUpCommandBar(ref p, "UpdateNames", RangeReferences.activeTable.UpdateNames);
                CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "UpdateNames", RangeReferences.activeTable.UpdateNames);
            }
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "UpdateAllNames", Main.instance.references.UpdateAllNames);
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "UpdateAllDBNames", Main.instance.references.UpdateAllDBNames);
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "UpdateAllColors", Main.instance.references.UpdateAllColors);
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "UpdateAllPSFormulas", Main.instance.references.UpdateAllPSFormulas);
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "UpdateAllDBFormulas", Main.instance.references.UpdateAllDBFormulas);
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "UpdateAllParents", Main.instance.references.UpdateAllParents, true);
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "UpdateAllReferencesPS", Main.instance.references.UpdateAllReferencesPS);
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "UpdateAllReferencesDB", Main.instance.references.UpdateAllReferencesDB);
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "UpdateAllLevels", Main.instance.references.UpdateAllLevels, true);
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "CheckAllRanges", Main.instance.references.CheckAllRanges);
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "ShowAllFormulas", () => {
                Thread t = new Thread(() =>
                {
                    AllFormulas form = new AllFormulas();
                    form.FormClosed += (s, args) =>
                    {
                        System.Windows.Forms.Application.ExitThread();
                    };
                    form.Show();
                    System.Windows.Forms.Application.Run();
                });
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
            });
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "CheckDublicate", () =>
            {
                Main.instance.references.CheckDublicates();
                //ReferenceObject co = Main.instance.references.references.Values.Where(c => c != RangeReferences.activeTable && c.HasRangePS(RangeReferences.activeTable.PS.Range) == true).FirstOrDefault();

                //if (co != null)
                //{
                //    MessageBox.Show(co._name + " " + RangeReferences.activeTable._name);
                //}
            });
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "UpdateHeadParents", Main.instance.heads.UpdateParents);
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "CheckFormulas", () => {
                foreach (string s in Main.instance.formulas.formulas.Keys)
                {
                    if (!RangeReferences.idDictionary.ContainsKey(s))
                    {
                        GlobalMethods.Err("idDictionary not found: " + s);
                    }
                }
            });
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "testWriteMeters", () =>
            {
                CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                string format = "dd MMMM yyyy";
                string data = "2023-02-" + this.textBox1.Text;
                data = this.textBox1.Text.PadLeft(2, '0') + " " + this.lblMonth.Text + " " + this.lblYear.Text;
                DateTime result;
                DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                if (string.IsNullOrEmpty(this.textBox1.Text) || (Int32.Parse(this.textBox1.Text) <= 0 && Int32.Parse(this.textBox1.Text) > 31))
                {
                    MessageBox.Show("Не введена дата записи!");
                    return;
                }
                MetersWriterTest.WriteMeters1(result);
            });
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "Обновить счетчики", () =>
            {
                Main.instance.references.UpdateMeterAllDB();
            });
            CustomCellMenu.AddButtonToPopUpCommandBar(ref p, "Test", () => {
                EmcosWrite(new DateTime(2024, 06, 01), new DateTime(2024, 06, 19));
            });*/

        }
        public void OpenForm()
        {
            Thread t = new Thread(() =>
            {
                FormulaEditor form = new FormulaEditor(ref RangeReferences.activeTable, RangeReferences.ActiveL1);
                form.FormClosed += (s, args) => 
                { 
                    System.Windows.Forms.Application.ExitThread(); 
                };
                form.Show();
                System.Windows.Forms.Application.Run();
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }
        
        public void ResetContextMenu()
        {
            //foreach (int index in Main.menuIndexes)
            //{
            //    CustomCellMenu.cb = Main.instance.xlApp.CommandBars[index];
            //    CustomCellMenu.cb.Reset();
            //}
        }
        public void MyContextMenu()
        {
            
            foreach (CommandBarControl item in cb.Controls)
            {
                item.Visible = false;
            }

            if (selectedButtons.Contains("Копировать")) cb.FindControl(Tag:"Копировать").Visible = true;
            if (selectedButtons.Contains("Вставить")) cb.FindControl(Tag: "Вставить").Visible = true;
            if (selectedButtons.Contains("GoTo DB")) cb.FindControl(Tag: "GoTo DB").Visible = true;
            if (selectedButtons.Contains("Выделить")) cb.FindControl(Tag: "Выделить").Visible = true;
            if (selectedButtons.Contains("Переместить субъект")) cb.FindControl(Tag: "Переместить субъект").Visible = true;
            if (selectedButtons.Contains("Переименовать")) cb.FindControl(Tag: "Переименовать").Visible = true;
            if (selectedButtons.Contains("Переименовать head")) cb.FindControl(Tag: "Переименовать").Visible = true;
            if (selectedButtons.Contains("Добавить код для макетТЭП")) cb.FindControl(Tag: "Добавить в макетТЭП").Visible = true;
            if (selectedButtons.Contains("Изменить код для макетТЭП")) cb.FindControl(Tag: "Изменить код макетТЭП").Visible = true;
            if (selectedButtons.Contains("Удалить код для макетТЭП")) cb.FindControl(Tag: "Удалить из макетТЭП").Visible = true;

            if (selectedButtons.Contains("Добавить код для ТЭП")) cb.FindControl(Tag: "Добавить в ТЭП").Visible = true;
            if (selectedButtons.Contains("Изменить код для ТЭП")) cb.FindControl(Tag: "Изменить код ТЭП").Visible = true;
            if (selectedButtons.Contains("Удалить код для ТЭП")) cb.FindControl(Tag: "Удалить из ТЭП").Visible = true;

            if (selectedButtons.Contains("Добавить новый L1")) 
            {
                CommandBarPopup p = (CommandBarPopup)cb.FindControl(Tag: "AddNewL1");
                p.Visible = true;
                ClearPopupMenu(p);

                foreach (string n in Main.instance.colors.main.Keys)
                {
                    if (n != "subject" && n != "план" && !RangeReferences.activeTable.DB.HasItem(n))
                    {
                        CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
                        {
                            RangeReferences.activeTable.AddNewDBL1StandartOther(n);
                            RangeReferences.activeTable.AddNewPS(n, "ручное");
                            RangeReferences.activeTable.PS.childs[n].childs["ручное"].ChangeCod();
                        };
                        CommandBarButton b = (CommandBarButton)p.Controls.Add(Type: 1, Temporary: true);
                        b.Caption = n;
                        b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
                    }
                }
                if (p.accChildCount == 0)
                    p.Visible = false;
            }
            if (selectedButtons.Contains("Добавить новый L2"))
            {
                CommandBarPopup p = (CommandBarPopup)cb.FindControl(Tag: "AddNewL2");
                p.Visible = true;
                ClearPopupMenu(p);
                foreach (string n in Main.instance.colors.mainTitle.Keys)
                {
                    if (n != "по плану" && n != "по счетчику" && n != "счетчик" && n != "утвержденный" && n != "корректировка" && n != "заявка" && !RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].HasItem(n))
                    {
                        AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.DB.AddNewRange, RangeReferences.ActiveL1, n);
                    }
                }
                if (p.accChildCount == 0)
                    p.Visible = false;
            }
            if (selectedButtons.Contains("Выбрать из EMCOS")) cb.FindControl(Tag: "Выбрать из EMCOS").Visible = true;
            if (selectedButtons.Contains("Записать данные из EMCOS (Все дни)")) cb.FindControl(Tag: "Записать из EMCOS (Все дни)").Visible = true;
            if (selectedButtons.Contains("Очистить данные из EMCOS (Все дни)")) cb.FindControl(Tag: "Очистить из EMCOS (Все дни)").Visible = true;
            if (selectedButtons.Contains("Записать из EMCOS")) cb.FindControl(Tag: "Записать данные из EMCOS").Visible = true;
            if (selectedButtons.Contains("Очистить из EMCOS")) cb.FindControl(Tag: "Очистить данные из EMCOS").Visible = true;
            if (selectedButtons.Contains("Изменить из EMCOS")) cb.FindControl(Tag: "Изменить из EMCOS").Visible = true;
            if (selectedButtons.Contains("Удалить из EMCOS")) cb.FindControl(Tag: "Удалить из EMCOS").Visible = true;
            if (selectedButtons.Contains("Удалить")) 
            {
                CommandBarPopup p = (CommandBarPopup)cb.FindControl(Tag: "RemoveOld");
                p.Visible = true;
                ClearPopupMenu(p);
                foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
                {
                    if (n != "код" && n != "основное" && n != "по плану" && n != "корректировка факт" && n != "ручное" && n != "счетчик" && n != "по счетчику")
                    {
                        //AddButtonToPopUpCommandBar(ref p, n, new List<Action>()
                        CustomCellMenu.AddButtonToPopUpCommandBar(ref p, n, new List<Action>()
                        {
                            Main.instance.StopAll,
                            () =>
                            {
                                if (RangeReferences.activeTable.PS.HasItem(n, SymbolType.lower))
                                {
                                    RangeReferences.activeTable.ChangeType(n, "ручное");
                                }
                                if (RangeReferences.activeTable.PS.HasItem(n, SymbolType.uper))
                                {
                                    RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].childs[n.ToUpper()].Remove();
                                }
                                if (n == "аскуэ")
                                {
                                    RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].emcosID = null;
                                    RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].emcosMLID = null;
                                }
                            },
                            () =>
                            {
                                RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs[n].Remove(false);
                            },
                            Main.instance.ResumeAll,
                            RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].ChangeCod
                        });
                    }
                }
                if (p.accChildCount == 0)
                    p.Visible = false;
            }
            if (selectedButtons.Contains("Показать")) 
            {
                CommandBarPopup p = (CommandBarPopup)cb.FindControl(Tag: "ShowTypeMenu");
                p.Visible = true;
                ClearPopupMenu(p);
                foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
                {
                    if (n != "код" && n != "основное" && n != "по плану" && !RangeReferences.activeTable._activeChild._activeChild.HasItem(n, SymbolType.uper))
                    {
                        AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable._activeChild.AddNewRange, RangeReferences.ActiveL1, n.ToUpper());
                    }
                }
                if (p.accChildCount == 0)
                    p.Visible = false;
            }
            if (selectedButtons.Contains("Скрыть")) cb.FindControl(Tag: "Скрыть").Visible = true;
            if (selectedButtons.Contains("Изменить main")) 
            {
                CommandBarPopup p = (CommandBarPopup)cb.FindControl(Tag: "ChangeTypeMenuMain");
                p.Visible = true;
                ClearPopupMenu(p);
                foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
                {
                    if (n != "корректировка факт" && n != "код" && n != "основное" && n != "счетчик" && !RangeReferences.activeTable._activeChild._activeChild.HasItem(n, SymbolType.lower))
                    {
                        AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.ChangeType, RangeReferences.activeL2, n);
                    }
                }

                if (p.accChildCount == 0)
                    p.Visible = false;
            }
            if (selectedButtons.Contains("Изменить mainSubtitle")) 
            {
                CommandBarPopup p = (CommandBarPopup)cb.FindControl(Tag: "ChangeTypeCellMenuMain");
                p.Visible = true;
                ClearPopupMenu(p);
                foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
                {
                    if (n != "корректировка факт" && n != "код" && n != "основное" && n != "счетчик" && !RangeReferences.activeTable._activeChild._activeChild.HasItem(n, SymbolType.lower))
                    {
                        AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.ChangeTypeCell, RangeReferences.activeL2, n);
                    }
                }

                if (p.accChildCount == 0)
                    p.Visible = false;
            }
            if (selectedButtons.Contains("Изменить extra")) 
            {
                CommandBarPopup p = (CommandBarPopup)cb.FindControl(Tag: "ChangeTypeMenuExtra");
                p.Visible = true;
                ClearPopupMenu(p);
                foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
                {
                    if (n != "код" && n != "основное" && n != "по плану")
                    {
                        AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.ChangeType, RangeReferences.activeL2, n.ToUpper());
                        //AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.ChangeType, RangeReferences.activeL2, n.ToUpper());
                    }
                }

                if (p.accChildCount == 0)
                    p.Visible = false;
            }
            if (selectedButtons.Contains("Ввести корректировку")) cb.FindControl(Tag: "Ввести корректировку").Visible = true;
            if (selectedButtons.Contains("Добавить план")) cb.FindControl(Tag: "Добавить план").Visible = true;
            if (selectedButtons.Contains("Изменить код плана")) cb.FindControl(Tag: "Изменить код плана").Visible = true;
            if (selectedButtons.Contains("Удалить план")) cb.FindControl(Tag: "Удалить план").Visible = true;
            if (selectedButtons.Contains("Изменить формулу")) cb.FindControl(Tag: "Изменить формулу").Visible = true;
            if (selectedButtons.Contains("Зависимые формулы")) cb.FindControl(Tag: "Зависимые формулы").Visible = true;
            if (selectedButtons.Contains("Добавить по показаниям счетчика")) cb.FindControl(Tag: "Добавить по показаниям счетчика").Visible = true;
            if (selectedButtons.Contains("Ввести показания счетчика")) cb.FindControl(Tag: "Ввести показания счетчика").Visible = true;
            if (selectedButtons.Contains("Изменить коэффициент счетчика")) cb.FindControl(Tag: "Изменить коэффициент счетчика").Visible = true;
            if (selectedButtons.Contains("Удалить по показаниям счетчика")) cb.FindControl(Tag: "Удалить по показаниям счетчика").Visible = true;
            if (selectedButtons.Contains("Сбросить")) cb.FindControl(Tag: "Сбросить").Visible = true;
            if (selectedButtons.Contains("Сбросить mainSubtitle")) cb.FindControl(Tag: "Сбросить").Visible = true;
            if (selectedButtons.Contains("Special")) SpecialMenuMain();
            if (selectedButtons.Contains("Удалить субъект")) cb.FindControl(Tag: "Удалить субъект").Visible = true;
            if (selectedButtons.Contains("Удалить тип")) cb.FindControl(Tag: "Удалить тип").Visible = true;

            if (selectedButtons.Contains("Добавить отступ")) cb.FindControl(Tag: "Добавить отступ").Visible = true;
            if (selectedButtons.Contains("Удалить отступ")) cb.FindControl(Tag: "Удалить отступ").Visible = true;
            if (selectedButtons.Contains("Удалить head")) cb.FindControl(Tag: "Удалить").Visible = true;
        }
        public void ClearContextMenu()
        {
            foreach (CommandBarControl item in CustomCellMenu.cb.Controls)
            {
                item.Visible = false;
            }
            //foreach (int index in Main.menuIndexes)
            //{
            //    //cb = Main.instance.xlApp.CommandBars[Main.menuIndexes[0]];
            //    //CustomCellMenu.cb = Main.instance.xlApp.CommandBars[index];
            //    foreach (CommandBarControl item in CustomCellMenu.cb.Controls)
            //    {
            //        item.Delete();
            //    }
            //}
        }
    }
}

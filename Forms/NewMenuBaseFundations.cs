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

        protected delegate void CommandBarButtonClick(CommandBarButton commandBarButton, ref bool cancel);
        private static object[,] oldValsArray;
        private static string oldVal;
        protected static string token = "";
        protected CommandBar cb;
        protected Color activeColor;
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
    
        #region AddButtonToCommandBar
        private void ContextMenuClickLog(string caption)
        {
            GlobalMethods.ToLog("Нажат пункт контекстного меню '" + caption + "'");
        }
        protected void AddButtonToCommandBar(string caption, Action action, int? faceid = null, int type = 1)
        {
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(caption);
                action.Invoke();
            };
            CommandBarButton b = (CommandBarButton)cb.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            if (faceid != null) b.FaceId = (int)faceid;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        protected void AddButtonToCommandBar(string caption, Action<string> action, string s1, int type = 1)
        {
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(caption);
                action.Invoke(s1);
            };
            CommandBarButton b = (CommandBarButton)cb.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        protected void AddButtonToCommandBar(string caption, Action<string, string> action, string s1, string s2, int type = 1)
        {
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(caption);
                action.Invoke(s1, s2);
            };
            CommandBarButton b = (CommandBarButton)cb.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }

        public void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, Action action, int type = 1)
        {
            string pName = p.Caption;
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(pName + " => " + caption);
                action.Invoke();
            };
            CommandBarButton b = (CommandBarButton)p.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        public void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, Action<bool> action, bool b1, int type = 1)
        {
            string pName = p.Caption;
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(pName + " => " + caption);
                action.Invoke(b1);
            };
            CommandBarButton b = (CommandBarButton)p.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        public void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, Action<string> action, string s1, int type = 1)
        {
            string pName = p.Caption;
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(pName + " => " + caption);
                action.Invoke(s1);
            };
            CommandBarButton b = (CommandBarButton)p.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        public void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, Action<string, string> action, string s1, string s2, int type = 1)
        {
            string pName = p.Caption;
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(pName + " => " + caption);
                action.Invoke(s1, s2);
            };
            CommandBarButton b = (CommandBarButton)p.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        public void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, Action<string, string, bool> action, string s1, string s2, bool b1 = true, int type = 1)
        {
            string pName = p.Caption;
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(pName + " => " + caption);
                action.Invoke(s1, s2, b1);
            };
            CommandBarButton b = (CommandBarButton)p.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        public void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, Action<string, string, string?> action, string s1, string s2, string? s3 = null, int type = 1)
        {
            string pName = p.Caption;
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(pName + " => " + caption);
                action.Invoke(s1, s2, s3);
            };
            CommandBarButton b = (CommandBarButton)p.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        public void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, List<Action> action, int type = 1)
        {
            string pName = p.Caption;
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(pName + " => " + caption);
                foreach (Action item in action)
                {
                    item.Invoke();
                }
            };
            CommandBarButton b = (CommandBarButton)p.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);   
        }
        public void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, List<Action<string>> action, string s1, int type = 1)
        {
            string pName = p.Caption;
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(pName + " => " + caption);
                foreach (Action<string> item in action)
                {
                    item.Invoke(s1);
                }
            };
            CommandBarButton b = (CommandBarButton)p.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        #endregion
    
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
            Marshal.ReleaseComObject(cb);
            if (_activeRange != null) Marshal.ReleaseComObject(_activeRange);
        }
        
        protected virtual void GotoDB()
        {
            //MessageBox.Show(RangeReferences.activeTable._name);
            Main.instance.wsDb.Activate();
            RangeReferences.activeTable.DB.Range.Select();
        }
        protected virtual void SelectSubject()
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
        private void ChangeTypeMenuMain()
        {
            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Изменить";
            foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
            {
                if (n != "корректировка факт" && n != "код" && n != "основное" && n != "счетчик" && !RangeReferences.activeTable._activeChild._activeChild.HasItem(n, SymbolType.lower))
                {
                    AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.ChangeType, RangeReferences.activeL2, n);
                }
            }

            if (p.accChildCount == 0)
                p.Delete();
        }
        private void ChangeTypeCellMenuMain()
        {
            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Изменить";
            foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
            {
                if (n != "корректировка факт" && n != "код" && n != "основное" && n != "счетчик" && !RangeReferences.activeTable._activeChild._activeChild.HasItem(n, SymbolType.lower))
                {
                    AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.ChangeTypeCell, RangeReferences.activeL2, n);
                }
            }

            if (p.accChildCount == 0)
                p.Delete();
        }
        private void ChangeTypeMenuExtra()
        {
            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Изменить";
            foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
            {
                if (n != "код" && n != "основное" && n != "по плану")
                {
                    AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.ChangeType, RangeReferences.activeL2, n.ToUpper());
                }
            }

            if (p.accChildCount == 0)
                p.Delete();
        }
        private void ShowTypeMenu()
        {
            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Показать";
            foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
            {
                if (n != "код" && n != "основное" && n != "по плану" && !RangeReferences.activeTable._activeChild._activeChild.HasItem(n, SymbolType.uper))
                {
                    AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable._activeChild.AddNewRange, RangeReferences.ActiveL1, n.ToUpper());
                }
            }
            if (p.accChildCount == 0)
                p.Delete();
        }
        private void HideType()
        {
            RangeReferences.activeTable._activeChild._activeChild._activeChild.Remove();
        }
        protected void AddNewL1()
        {
            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Добавить новый";
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
                p.Delete();
        }
        protected void AddNewL2()
        {
            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Добавить новый";
            foreach (string n in Main.instance.colors.mainTitle.Keys)
            {
                if (n != "по плану" && n != "по счетчику" && n != "счетчик" && n != "утвержденный" && n != "корректировка" && n != "заявка" && !RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].HasItem(n))
                {
                    AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.DB.AddNewRange, RangeReferences.ActiveL1, n);
                }
            }
            if (p.accChildCount == 0)
                p.Delete();
        }
        protected void EmcosSelect()
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
        protected void EmcosRemove()
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
        public void RemoveOld()
        {
            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Удалить";
            foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
            {
                if (n != "код" && n != "основное" && n != "по плану" && n != "корректировка факт" && n != "ручное" && n != "счетчик" && n != "по счетчику")
                {                        
                    AddButtonToPopUpCommandBar(ref p, n, new List<Action>()
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
                p.Delete();
        }
        protected void SpecialMenuMain()
        {
            Action action;

            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Special";
            if (Main.instance.colors.main["subject"] == activeColor)
            {
                AddButtonToPopUpCommandBar(ref p, "UpdateNames", RangeReferences.activeTable.UpdateNames);
            }
            AddButtonToPopUpCommandBar(ref p, "UpdateAllNames", Main.instance.references.UpdateAllNames);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllDBNames", Main.instance.references.UpdateAllDBNames);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllColors", Main.instance.references.UpdateAllColors);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllPSFormulas", Main.instance.references.UpdateAllPSFormulas);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllDBFormulas", Main.instance.references.UpdateAllDBFormulas);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllParents", Main.instance.references.UpdateAllParents,  true);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllReferencesPS", Main.instance.references.UpdateAllReferencesPS);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllReferencesDB", Main.instance.references.UpdateAllReferencesDB);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllLevels", Main.instance.references.UpdateAllLevels, true);
            AddButtonToPopUpCommandBar(ref p, "CheckAllRanges", Main.instance.references.CheckAllRanges);
            AddButtonToPopUpCommandBar(ref p, "ShowAllFormulas",() => {
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
            AddButtonToPopUpCommandBar(ref p, "CheckDublicate", () => 
            {
                Main.instance.references.CheckDublicates();
                //ReferenceObject co = Main.instance.references.references.Values.Where(c => c != RangeReferences.activeTable && c.HasRangePS(RangeReferences.activeTable.PS.Range) == true).FirstOrDefault();

                //if (co != null)
                //{
                //    MessageBox.Show(co._name + " " + RangeReferences.activeTable._name);
                //}
            });
            AddButtonToPopUpCommandBar(ref p, "UpdateHeadParents", Main.instance.heads.UpdateParents);
            AddButtonToPopUpCommandBar(ref p, "CheckFormulas", () => {
                foreach (string s in Main.instance.formulas.formulas.Keys)
                {
                    if (!RangeReferences.idDictionary.ContainsKey(s))
                    {
                        GlobalMethods.Err("idDictionary not found: " + s);
                    }
                }
            });
            AddButtonToPopUpCommandBar(ref p, "testWriteMeters", () => 
            {
                CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                string format = "dd MMMM yyyy";
                string data = "2023-02-" + this.textBox1.Text;
                data = this.textBox1.Text.PadLeft(2,'0') + " " + this.lblMonth.Text + " " + this.lblYear.Text;
                DateTime result;
                DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                if (string.IsNullOrEmpty(this.textBox1.Text) ||  (Int32.Parse(this.textBox1.Text) <= 0 && Int32.Parse(this.textBox1.Text) > 31))
                {
                    MessageBox.Show("Не введена дата записи!");
                    return;
                }
                MetersWriterTest.WriteMeters1(result);
            });
            AddButtonToPopUpCommandBar(ref p, "Обновить счетчики", () => 
            {
                Main.instance.references.UpdateMeterAllDB();
            });
            AddButtonToPopUpCommandBar(ref p, "Test", ()=>{
                EmcosWrite(new DateTime(2024, 06, 01), new DateTime(2024, 06, 19));
            });
        }
        protected void OpenForm()
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
            foreach (int index in Main.menuIndexes)
            {
                cb = Main.instance.xlApp.CommandBars[index];
                cb.Reset();
            }
        }
        public void MyContextMenu()
        {
            foreach (int index in Main.menuIndexes)
            {
                //cb = Main.instance.xlApp.CommandBars[38];
                cb = Main.instance.xlApp.CommandBars[index];
                foreach (CommandBarControl item in cb.Controls)
                {
                    item.Delete();
                    Marshal.ReleaseComObject(item);
                }

                if (selectedButtons.Contains("Копировать")) AddButtonToCommandBar("Копировать", () => 
                {
                    GlobalMethods.ToLog("Копирование диапазона: " + ((Excel.Range)Main.instance.xlApp.Selection).Address);
                    ((Excel.Range)Main.instance.xlApp.Selection).Copy();
                }, 0019);
                if (selectedButtons.Contains("Вставить")) AddButtonToCommandBar("Вставить", () => 
                {
                    try
                    {
                        IDataObject idat = null;
                        Invoke(new Action(() => idat = Clipboard.GetDataObject()));
                        if (idat != null)
                        {
                            if (idat.GetDataPresent(DataFormats.Text))
                            {
                                string clipboardString = idat.GetData(DataFormats.Text) as string;
                                // clipboardString = clipboardString.Replace(",", ".");

                                string[] rows = clipboardString.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                                int numRows = rows.Length;
                                int numColumns = rows[0].Split('\t').Length;

                                Excel.Range selectedRange = (Excel.Range)Main.instance.xlApp.Selection;
                                int targetNumRows = selectedRange.Rows.Count;
                                int targetNumColumns = selectedRange.Columns.Count;

                                // Расширение выделенного диапазона, если необходимо
                                if (targetNumRows < numRows || targetNumColumns < numColumns)
                                {
                                    Excel.Range expandedRange = selectedRange.Resize[numRows, numColumns];
                                    selectedRange = expandedRange;
                                    targetNumRows = selectedRange.Rows.Count;
                                    targetNumColumns = selectedRange.Columns.Count;
                                }

                                object[,] dataArray = new object[targetNumRows, targetNumColumns];
                                for (int i = 0; i < targetNumRows; i++)
                                {
                                    for (int j = 0; j < targetNumColumns; j++)
                                    {
                                        int sourceRow = i % numRows;
                                        int sourceColumn = j % numColumns;
                                        double val;
                                        string strVal = rows[sourceRow].Split('\t')[sourceColumn];
                                        if (double.TryParse(strVal, out val))
                                        {
                                            dataArray[i, j] = val;
                                        }
                                        else
                                        {
                                            dataArray[i, j] = null;
                                        }
                                        
                                    }
                                }
                                selectedRange.Select();
                                selectedRange.Value2 = dataArray;
                                GlobalMethods.ToLog(" в диапазон: " + selectedRange.Address + "Вставлены значения: \n" + clipboardString );
                            }
                            else if (idat.GetDataPresent(DataFormats.StringFormat))
                            {
                                // string clipboardString = idat.GetData(DataFormats.StringFormat) as string;
                                // // clipboardString = clipboardString.Replace("\r\n", "");
                                // clipboardString = clipboardString.Replace(",", ".");
                                // Main.instance.xlApp.ActiveCell.Value = clipboardString;
                            }
                            else if (idat.GetDataPresent(DataFormats.UnicodeText))
                            {
                                // string clipboardString = idat.GetData(DataFormats.UnicodeText) as string;
                                // // clipboardString = clipboardString.Replace("\r\n", "");
                                // clipboardString = clipboardString.Replace(",", ".");
                                // Main.instance.xlApp.ActiveCell.Value = clipboardString;
                            }
                            else
                            {
                                MessageBox.Show("Ошибка вставки данных");
                            }
                        }
                    }
                    catch (Exception){}
                }, 0022);
                if (selectedButtons.Contains("GoTo DB")) AddButtonToCommandBar("GoTo DB", GotoDB, 2116);
                if (selectedButtons.Contains("Выделить")) AddButtonToCommandBar("Выделить", () => 
                {
                    SelectSubject();
                }, 118);
                if (selectedButtons.Contains("Переместить субъект")) AddButtonToCommandBar("Переместить субъект", () => 
                {
                    using (TransferSubject form = new TransferSubject())
                    {
                        form.ShowDialog();
                    }
                });
                if (selectedButtons.Contains("Переименовать")) AddButtonToCommandBar("Переименовать", () => 
                {
                    using (Rename form = new Rename(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                }, 7677);
                if (selectedButtons.Contains("Переименовать head")) AddButtonToCommandBar("Переименовать", () => 
                {
                    HeadObject ho = Main.instance.heads.HeadByRange(_activeRange);
                    if (ho != null)
                    {
                        using (Rename form = new Rename(ho))
                        {
                            form.ShowDialog();
                        }
                    }
                }, 7677);
                if (selectedButtons.Contains("Добавить код для макетТЭП")) AddButtonToCommandBar("Добавить в макетТЭП", () => 
                {
                    ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                    using (AddPlan form = new AddPlan(co))
                    {
                        form.ShowDialog();
                    }
                });
                if (selectedButtons.Contains("Изменить код для макетТЭП")) AddButtonToCommandBar("Изменить код макетТЭП", () => 
                {
                    ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                    using (AddPlan form = new AddPlan(co))
                    {
                        form.ShowDialog();
                    }
                });
                if (selectedButtons.Contains("Удалить код для макетТЭП")) AddButtonToCommandBar("Удалить из макетТЭП", () => 
                {
                    RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].RemoveFromMTEP();
                    // int cod = (int)RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].codMaketTEP;
                    // Main.instance.wsMTEP.Range["A:A"].Find(cod).Interior.ColorIndex = 0;
                    // Main.instance.wsMTEP.Range["A:A"].Find(cod).Offset[0, 2].Value = "";
                    // RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].codMaketTEP = null;
                });

                if (selectedButtons.Contains("Добавить код для ТЭП")) AddButtonToCommandBar("Добавить в ТЭП", () =>
                {
                    ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                    using (AddTEP form = new AddTEP(co))
                    {
                        form.ShowDialog();
                    }
                });
                if (selectedButtons.Contains("Изменить код для ТЭП")) AddButtonToCommandBar("Изменить код ТЭП", () =>
                {
                    ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                    using (AddTEP form = new AddTEP(co))
                    {
                        form.ShowDialog();
                    }
                });
                if (selectedButtons.Contains("Удалить код для ТЭП")) AddButtonToCommandBar("Удалить из ТЭП", () =>
                {
                    RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].RemoveFromTEP();
                    // int cod = (int)RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].codTEP;
                    // string adr1 = Main.instance.wsTEPm.Range["5:5"].Find(What: cod, LookAt: Excel.XlLookAt.xlWhole).Address[false, false];
                    // string adr2 = Main.instance.wsTEPm.Range["5:5"].Find(What: cod, LookAt: Excel.XlLookAt.xlWhole).Offset[0, 1].Address[false, false];
                    // string adr = Regex.Replace(adr1, @"[^A-Z]+", String.Empty) + ":" + Regex.Replace(adr2, @"[^A-Z]+", String.Empty);

                    // Main.instance.StopAll();
                    // Main.instance.wsTEPn.Range[adr].Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
                    // Main.instance.wsTEPm.Range[adr].Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
                    // Main.instance.ResumeAll();

                    // RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].codTEP = null;
                });
                
                if (selectedButtons.Contains("Добавить новый L1")) AddNewL1();
                if (selectedButtons.Contains("Добавить новый L2")) AddNewL2();
                if (selectedButtons.Contains("Выбрать из EMCOS")) AddButtonToCommandBar("Выбрать из EMCOS", EmcosSelect);
                if (selectedButtons.Contains("Записать данные из EMCOS (Все дни)")) AddButtonToCommandBar("Записать из EMCOS (Все дни)", () => 
                {
                    CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                    string format = "dd MMMM yyyy";
                    string data = "01" + " " + this.lblMonth.Text + " " + this.lblYear.Text;
                    DateTime result;
                    DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                    EmcosWrite(result, DateTime.Today.AddDays(-1), RangeReferences.activeTable);
                });
                if (selectedButtons.Contains("Очистить данные из EMCOS (Все дни)")) AddButtonToCommandBar("Очистить из EMCOS (Все дни)", () => 
                {
                    CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                    string format = "dd MMMM yyyy";
                    string data = "01" + " " + this.lblMonth.Text + " " + this.lblYear.Text;
                    DateTime result;
                    DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                    EmcosClear(result, DateTime.Today.AddDays(-1), RangeReferences.activeTable);
                });
                if (selectedButtons.Contains("Записать из EMCOS")) AddButtonToCommandBar("Записать данные из EMCOS", () => 
                {
                    CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                    string format = "dd MMMM yyyy";
                    string data = RangeReferences.activeTable.ActiveDay().ToString().PadLeft(2, '0') + " " + this.lblMonth.Text + " " + this.lblYear.Text;
                    DateTime result;
                    DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                    EmcosWrite(result, result, RangeReferences.activeTable);
                });
                if (selectedButtons.Contains("Очистить из EMCOS")) AddButtonToCommandBar("Очистить данные из EMCOS", () => 
                {
                    CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                    string format = "dd MMMM yyyy";
                    string data = RangeReferences.activeTable.ActiveDay().ToString().PadLeft(2, '0') + " " + this.lblMonth.Text + " " + this.lblYear.Text;
                    DateTime result;
                    DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                    EmcosClear(result, result, RangeReferences.activeTable);
                });
                if (selectedButtons.Contains("Изменить из EMCOS")) AddButtonToCommandBar("Изменить из EMCOS", EmcosSelect);
                if (selectedButtons.Contains("Удалить из EMCOS")) AddButtonToCommandBar("Удалить из EMCOS", EmcosRemove);
                if (selectedButtons.Contains("Удалить")) RemoveOld();
                if (selectedButtons.Contains("Показать")) ShowTypeMenu();
                if (selectedButtons.Contains("Скрыть")) AddButtonToCommandBar("Скрыть", HideType);
                if (selectedButtons.Contains("Изменить main")) ChangeTypeMenuMain();
                if (selectedButtons.Contains("Изменить mainSubtitle")) ChangeTypeCellMenuMain();
                if (selectedButtons.Contains("Изменить extra")) ChangeTypeMenuExtra();
                if (selectedButtons.Contains("Ввести корректировку")) AddButtonToCommandBar("Ввести корректировку", () =>
                {
                    using(Correct form = new Correct(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                }, 387);
                if (selectedButtons.Contains("Добавить план")) AddButtonToCommandBar("Добавить план", () => 
                {
                    using (AddPlan form = new AddPlan(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                }, 213);
                if (selectedButtons.Contains("Изменить код плана")) AddButtonToCommandBar("Изменить код плана", () => 
                {
                    using (AddPlan form = new AddPlan(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                }, 712);
                if (selectedButtons.Contains("Удалить план")) AddButtonToCommandBar("Удалить план", () => 
                {
                    RangeReferences.activeTable.codPlan = null; 
                    RangeReferences.activeTable.RemovePlan();
                }, 214);
                if (selectedButtons.Contains("Изменить формулу")) AddButtonToCommandBar("Изменить формулу", () =>
                {
                    OpenForm();
                }, 385);
                if (selectedButtons.Contains("Зависимые формулы")) AddButtonToCommandBar("Зависимые формулы", () => 
                {
                    Thread t = new Thread(() =>
                    {
                        AllFormulas form = new AllFormulas(RangeReferences.activeTable, RangeReferences.ActiveL1);
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
                if (selectedButtons.Contains("Добавить по показаниям счетчика")) AddButtonToCommandBar("Добавить по показаниям счетчика",() => {
                    RangeReferences.activeTable.AddMeter(RangeReferences.ActiveL1);
                    }, 33);
                if (selectedButtons.Contains("Ввести показания счетчика")) AddButtonToCommandBar("Ввести показания счетчика", () =>
                    {
                        using (SCH form = new SCH(RangeReferences.activeTable))
                        {
                            form.ShowDialog();
                        }
                    }, 205);
                if (selectedButtons.Contains("Изменить коэффициент счетчика")) AddButtonToCommandBar("Изменить коэффициент счетчика", () =>
                    {
                        using (ChangeCoef form = new ChangeCoef(RangeReferences.activeTable))
                        {
                            form.ShowDialog();
                        }
                    }, 400);
                if (selectedButtons.Contains("Удалить по показаниям счетчика")) AddButtonToCommandBar("Удалить по показаниям счетчика", () => { 
                    RangeReferences.activeTable.RemoveMeter(RangeReferences.ActiveL1); 
                    });
                if (selectedButtons.Contains("Сбросить")) AddButtonToCommandBar("Сбросить", () => {
                    RangeReferences.activeTable.Reset(RangeReferences.ActiveL1, RangeReferences.activeL2);
                });
                if (selectedButtons.Contains("Сбросить mainSubtitle")) AddButtonToCommandBar("Сбросить", () => {
                    RangeReferences.activeTable.ResetCell(RangeReferences.ActiveL1, RangeReferences.activeL2);
                });
                if (selectedButtons.Contains("Special")) SpecialMenuMain();
                if (selectedButtons.Contains("Удалить субъект"))AddButtonToCommandBar("Удалить субъект", () => RangeReferences.activeTable.RemoveSubject(), 330);
                if (selectedButtons.Contains("Удалить тип"))AddButtonToCommandBar("Удалить тип", () => RangeReferences.activeTable.RemoveChild(RangeReferences.ActiveL1));

                if (selectedButtons.Contains("Добавить отступ")) AddButtonToCommandBar("Добавить отступ", () =>
                {
                    HeadObject ho = Main.instance.heads.HeadByRange(_activeRange);
                    if (ho != null)
                    {
                        ho.Indent(IndentDirection.right);
                    }
                }, faceid: 137);
                if (selectedButtons.Contains("Удалить отступ")) AddButtonToCommandBar("Удалить отступ", () =>
                {
                    HeadObject ho = Main.instance.heads.HeadByRange(_activeRange);
                    if (ho != null)
                    {
                        ho.Indent(IndentDirection.right);
                    }
                }, faceid: 138);
                if (selectedButtons.Contains("Удалить head")) AddButtonToCommandBar("Удалить", () => 
                {
                    HeadObject ho = Main.instance.heads.HeadByRange(_activeRange);
                    if (ho != null)
                    {
                        if (MessageBox.Show("Это удалит всех субъектов входящих в " + ho._name + "\nВы Уверены?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                        {
                            ho.Delete();
                        }
                    }
                }, 1088);
            }
        }
        public void ClearContextMenu()
        {
            foreach (int index in Main.menuIndexes)
            {
                //cb = Main.instance.xlApp.CommandBars[Main.menuIndexes[0]];
                cb = Main.instance.xlApp.CommandBars[index];
                foreach (CommandBarControl item in cb.Controls)
                {
                    item.Delete();
                }
            }
        }
    }
}

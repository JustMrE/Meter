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
//using Microsoft.Office.Interop.Excel;

namespace Meter.Forms
{
    partial class NewMenuBase
    {
        [DllImport("user32.dll", SetLastError = true)]
        protected static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        protected static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOOLWINDOW = 0x80;
        public IntPtr formHwnd;

        protected delegate void CommandBarButtonClick(CommandBarButton commandBarButton, ref bool cancel);
        private static object[,] oldValsArray;
        private static string oldVal;
        //public static HttpClient client = new HttpClient();
        protected static string token = "";
        protected CommandBar cb;
        protected Color activeColor;
        public static Excel.Range _activeRange;
        public static bool restartFlag = false;
        public static HashSet<string> editedFormulas = new HashSet<string>();
        public static HashSet<string> selectedButtons = new HashSet<string>();

        
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
                        var matchingIds = deviceIds.Where(id => Regex.IsMatch(id._name, pattern: search, ro)).Select(n => n._name).ToArray();
                        listBox1.Items.AddRange(matchingIds.ToArray());
                    }
                    catch (ArgumentException)
                    {
                        this.Invoke((MethodInvoker)(() =>
                        {
                            listBox1.Items.Clear();
                            listBox1.Items.AddRange(Main.instance.references.references.Keys.OrderBy(m => m).ToArray());
                        }));
                    }

                }));
            }
            else
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    listBox1.Items.Clear();
                    listBox1.Items.AddRange(Main.instance.references.references.Keys.OrderBy(m => m).ToArray());
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
                Main.dontsave = true;
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
                selectedButtons.Add("Ввести корректировку");
                if (RangeReferences.activeTable.DB.HasItem("по счетчику"))
                {
                    selectedButtons.Add("Ввести показания счетчика");
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
                if (!RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].HasItem("счетчик"))
                {
                    selectedButtons.Add("Добавить по показаниям счетчика");
                }
                else
                {
                    selectedButtons.Add("Удалить по показаниям счетчика");
                }
            }
            if (activeColor == Main.instance.colors.main["subject"])
            {
                if (RangeReferences.activeTable.DB.HasItem("план"))
                {
                    selectedButtons.Add("Изменить код плана");
                    selectedButtons.Add("Удалить план");
                }
                else
                {
                    selectedButtons.Add("Добавить план");
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

            // if (!myMenuButtons.ContainsKey(Tag))
            // {
            //     myMenuButtons.Add(Tag, b);
            // }
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
            Main.instance.references.ActivateTable(range);
            if (RangeReferences._activeObject != null) RangeReferences._activeObject.Range.Select();
        }
        public virtual void SlectionChanged(Excel.Range range)
        {
            _activeRange = range;
            Main.instance.references.ActivateTable(range);
            if (range.Formula is string)
            {
                oldVal = (string)range.Formula;
            }
            else if (range.Formula is object[,])
            {
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
            {
                if (range.Cells.Count != _activeRange.Cells.Count)
                {
                    int r, c;
                    Excel.Range r1 = ((Excel.Range)_activeRange.Cells[1, 1]);

                    r = 1 + (range.Row - r1.Row);
                    c = 1 + (range.Column - r1.Column);
                    oldVal = (string)oldValsArray[r, c];
                    Marshal.ReleaseComObject(r1);
                }
                if (range.Cells.Count > 1) 
                {
                    DontEditRange(range);
                }
                else
                {
                    Color c = ColorsData.GetRangeColor(range);
                    if (c == Main.instance.colors.mainSubtitle["ручное"] || 
                        c == Main.instance.colors.mainSubtitle["корректировка"] || 
                        c == Main.instance.colors.mainSubtitle["корректировка факт"] ||
                        c == Main.instance.colors.mainSubtitle["счетчик"] ||
                        c == Main.instance.colors.extraSubtitle["РУЧНОЕ"] ||
                        c == Main.instance.colors.extraSubtitle["КОРРЕКТИРОВКА"] ||
                        c == Main.instance.colors.extraSubtitle["КОРРЕКТИРОВКА ФАКТ"] ||
                        c == Main.instance.colors.extraSubtitle["СЧЕТЧИК"])
                    {
                        int? day = RangeReferences.activeTable.ActiveDay();
                        if (day != null)
                        {
                            
                            WriteToDB(_activeRange, (int)day);
                        }
                    }
                    else
                    {
                        DontEdit(range);
                    }
                }
            }
            Main.instance.xlApp.EnableEvents = true;
        }
        public void WriteToDB(Excel.Range rng, int day)
        {
            string val = rng.Formula != null ? rng.Formula.ToString() : "";
            DontEdit(rng, day);
            RangeReferences.activeTable.WriteToDB(RangeReferences.ActiveL1, RangeReferences.activeL2,(int)day, val);
            
        }
        private void DontEdit(Excel.Range rng, int? day = null)
        {
            string val = rng.Formula != null ? rng.Formula.ToString() : "";
            rng.Formula = oldVal;
            GlobalMethods.ToLog("Изменено значение ячейки " + rng.Address + " с '" + val + "' на '" + oldVal + "'");
        }
        private void DontEditRange(Excel.Range rng)
        {
            object[,] newValsArray = (object[,])rng.Formula;
            for (int i = 1; i <= rng.Columns.Count; i++)
            {
                for (int j = 1; j <= rng.Rows.Count; j++)
                {
                    GlobalMethods.ToLog("Изменено значение ячейки " + ((Excel.Range)rng.Cells[j, i]).Address + " с '" + newValsArray[j, i] + "' на '" + oldValsArray[j, i] + "'");
                }
            }
            rng.Formula = oldValsArray;
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
        private void ChangeTypeMenuMain()
        {
            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Изменить";
            foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
            {
                if (n != "корректировка факт" && n != "код" && n != "основное" && !RangeReferences.activeTable._activeChild._activeChild.HasItem(n, SymbolType.lower))
                {
                    AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.ChangeType, RangeReferences.activeL2, n);
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
        protected void RemoveOld()
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
                        RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs[n].Remove,
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
            //AddButtonToPopUpCommandBar(ref p, "ClearAllDB", Main.instance.references.ClearAllDB);
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
            AddButtonToPopUpCommandBar(ref p, "DeleteHead", () => {
                HeadObject ho = Main.instance.heads.HeadByRange(_activeRange);
                if (ho != null)
                {
                    ho.Delete();
                }
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
            // foreach (int index in Main.menuIndexes)
            {
                cb = Main.instance.xlApp.CommandBars[Main.menuIndexes[0]];
                foreach (CommandBarControl item in cb.Controls)
                {
                    item.Delete();
                    Marshal.ReleaseComObject(item);
                }

                if (selectedButtons.Contains("GoTo DB")) AddButtonToCommandBar("GoTo DB", GotoDB);
                if (selectedButtons.Contains("Переименовать")) AddButtonToCommandBar("Переименовать", () => 
                {
                    using (Rename form = new Rename(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                });
                if (selectedButtons.Contains("Добавить новый L1")) AddNewL1();
                if (selectedButtons.Contains("Добавить новый L2")) AddNewL2();
                if (selectedButtons.Contains("Выбрать из EMCOS")) AddButtonToCommandBar("Выбрать из EMCOS", EmcosSelect);
                if (selectedButtons.Contains("Изменить из EMCOS")) AddButtonToCommandBar("Изменить из EMCOS", EmcosSelect);
                if (selectedButtons.Contains("Удалить из EMCOS")) AddButtonToCommandBar("Удалить из EMCOS", EmcosRemove);
                if (selectedButtons.Contains("Удалить")) RemoveOld();
                if (selectedButtons.Contains("Показать")) ShowTypeMenu();
                if (selectedButtons.Contains("Скрыть")) AddButtonToCommandBar("Скрыть", HideType);
                if (selectedButtons.Contains("Изменить main")) ChangeTypeMenuMain();
                if (selectedButtons.Contains("Изменить extra")) ChangeTypeMenuExtra();
                if (selectedButtons.Contains("Ввести корректировку")) AddButtonToCommandBar("Ввести корректировку", () =>
                {
                    using(Correct form = new Correct(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                });
                if (selectedButtons.Contains("Добавить план")) AddButtonToCommandBar("Добавить план", () => 
                {
                    using (AddPlan form = new AddPlan(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                });
                if (selectedButtons.Contains("Изменить код плана")) AddButtonToCommandBar("Изменить код плана", () => 
                {
                    using (AddPlan form = new AddPlan(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                });
                if (selectedButtons.Contains("Удалить план")) AddButtonToCommandBar("Удалить план", () => 
                {
                    RangeReferences.activeTable.codPlan = null; 
                    RangeReferences.activeTable.RemovePlan();
                });
                if (selectedButtons.Contains("Изменить формулу")) AddButtonToCommandBar("Изменить формулу", () =>
                {
                    OpenForm();
                });
                if (selectedButtons.Contains("Добавить по показаниям счетчика")) AddButtonToCommandBar("Добавить по показаниям счетчика",() => {
                    RangeReferences.activeTable.AddMeter(RangeReferences.ActiveL1);
                    });
                if (selectedButtons.Contains("Ввести показания счетчика")) AddButtonToCommandBar("Ввести показания счетчика", () =>
                    {
                        using (SCH form = new SCH(RangeReferences.activeTable))
                        {
                            form.ShowDialog();
                        }
                    });
                if (selectedButtons.Contains("Изменить коэффициент счетчика")) AddButtonToCommandBar("Изменить коэффициент счетчика", () =>
                    {
                        using (ChangeCoef form = new ChangeCoef(RangeReferences.activeTable))
                        {
                            form.ShowDialog();
                        }
                    });
                if (selectedButtons.Contains("Удалить по показаниям счетчика")) AddButtonToCommandBar("Удалить по показаниям счетчика", () => { 
                    RangeReferences.activeTable.RemoveMeter(RangeReferences.ActiveL1); 
                    });
                if (selectedButtons.Contains("Сбросить")) AddButtonToCommandBar("Сбросить", () => {
                    RangeReferences.activeTable.Reset(RangeReferences.ActiveL1, RangeReferences.activeL2);
                });
                if (selectedButtons.Contains("Special")) SpecialMenuMain();
                if (selectedButtons.Contains("Удалить субъект"))AddButtonToCommandBar("Удалить субъект", RangeReferences.activeTable.RemoveSubject);
                if (selectedButtons.Contains("Удалить тип"))AddButtonToCommandBar("Удалить тип", RangeReferences.activeTable._activeChild.Remove);

                if (selectedButtons.Contains("Добавить отступ")) AddButtonToCommandBar("Добавить отступ", () =>
                {
                    HeadObject ho = Main.instance.heads.HeadByRange(_activeRange);
                    if (ho != null)
                    {
                        ho.Indent(IndentDirection.right);
                    }
                }, faceid: 374);
                if (selectedButtons.Contains("Удалить отступ")) AddButtonToCommandBar("Удалить отступ", () =>
                {
                    HeadObject ho = Main.instance.heads.HeadByRange(_activeRange);
                    if (ho != null)
                    {
                        ho.Indent(IndentDirection.right);
                    }
                }, faceid: 375);
            }
        }
        public void ClearContextMenu()
        {
            // foreach (int index in Main.menuIndexes)
            {
                cb = Main.instance.xlApp.CommandBars[Main.menuIndexes[0]];
                foreach (CommandBarControl item in cb.Controls)
                {
                    item.Delete();
                }
            }
        }
    }
}

using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;
using Meter.Forms;
using System.Globalization;
using System.Windows.Forms;

namespace Meter
{
    public static class CustomCellMenu
    {
        public delegate void CommandBarButtonClick(CommandBarButton commandBarButton, ref bool cancel);

        public static CommandBar cb;
        public static void CreateCustomContextMenu()
        {
            try
            {
                // Получаем контекстное меню по умолчанию для ячеек
                var cellMenu = Main.instance.xlApp.CommandBars["Cell"];

                // Создаем новое пользовательское меню
                cb = Main.instance.xlApp.CommandBars.Add("CustomCellMenu", MsoBarPosition.msoBarPopup, false, true);

                AddButtonsToMenu();
                //// Добавляем пункт в пользовательское меню
                //var menuItem = (CommandBarButton)customContextMenu.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
                //menuItem.Caption = "My Custom Item";
                //menuItem.Click += new _CommandBarButtonEvents_ClickEventHandler(MenuItem_Click);

                //// Добавляем пользовательское меню в список меню ячейки
                //customContextMenu.Visible = true;
                //cellMenu.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error creating context menu: " + ex.Message);
            }
        }

        static void AddButtonsToMenu()
        {
            AddButtonToCommandBar("Копировать", () =>
            {
                GlobalMethods.ToLog("Копирование диапазона: " + ((Excel.Range)Main.instance.xlApp.Selection).Address);
                ((Excel.Range)Main.instance.xlApp.Selection).Copy();
            }, 0019);
            AddButtonToCommandBar("Вставить", () =>
            {
                try
                {
                    IDataObject idat = null;
                    Main.instance.menu.Invoke(new Action(() => idat = Clipboard.GetDataObject()));
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
                            GlobalMethods.ToLog(" в диапазон: " + selectedRange.Address + "Вставлены значения: \n" + clipboardString);
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
                catch (Exception) { }
            }, 0022);
            AddButtonToCommandBar("GoTo DB", Main.instance.menu.GotoDB, 2116);
            AddButtonToCommandBar("Выделить", () =>
            {
                Main.instance.menu.SelectSubject();
            }, 118);
            AddButtonToCommandBar("Переместить субъект", () =>
            {
                using (TransferSubject form = new TransferSubject())
                {
                    form.ShowDialog();
                }
            });
            AddButtonToCommandBar("Переименовать", () =>
            {
                using (Rename form = new Rename(RangeReferences.activeTable))
                {
                    form.ShowDialog();
                }
            }, 7677);
            AddButtonToCommandBar("Переименовать", () =>
            {
                HeadObject ho = Main.instance.heads.HeadByRange(NewMenuBase._activeRange);
                if (ho != null)
                {
                    using (Rename form = new Rename(ho))
                    {
                        form.ShowDialog();
                    }
                }
            }, 7677, tag: "Переименовать head");
            AddButtonToCommandBar("Добавить в макетТЭП", () =>
            {
                ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                using (AddPlan form = new AddPlan(co))
                {
                    form.ShowDialog();
                }
            });
            AddButtonToCommandBar("Изменить код макетТЭП", () =>
            {
                ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                using (AddPlan form = new AddPlan(co))
                {
                    form.ShowDialog();
                }
            });
            AddButtonToCommandBar("Удалить из макетТЭП", () =>
            {
                RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].RemoveFromMTEP();
                // int cod = (int)RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].codMaketTEP;
                // Main.instance.wsMTEP.Range["A:A"].Find(cod).Interior.ColorIndex = 0;
                // Main.instance.wsMTEP.Range["A:A"].Find(cod).Offset[0, 2].Value = "";
                // RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].codMaketTEP = null;
            });

            AddButtonToCommandBar("Добавить в ТЭП", () =>
            {
                ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                using (AddTEP form = new AddTEP(co))
                {
                    form.ShowDialog();
                }
            });
            AddButtonToCommandBar("Изменить код ТЭП", () =>
            {
                ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                using (AddTEP form = new AddTEP(co))
                {
                    form.ShowDialog();
                }
            });
            AddButtonToCommandBar("Удалить из ТЭП", () =>
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

            AddNewL1();
            AddNewL2();
            AddButtonToCommandBar("Выбрать из EMCOS", Main.instance.menu.EmcosSelect);
            AddButtonToCommandBar("Записать из EMCOS (Все дни)", () =>
            {
                CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                string format = "dd MMMM yyyy";
                string data = "01" + " " + Main.instance.menu.lblMonth.Text + " " + Main.instance.menu.lblYear.Text;
                DateTime result;
                DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                Main.instance.menu.EmcosWrite(result, DateTime.Today.AddDays(-1), RangeReferences.activeTable);
            });
            AddButtonToCommandBar("Очистить из EMCOS (Все дни)", () =>
            {
                CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                string format = "dd MMMM yyyy";
                string data = "01" + " " + Main.instance.menu.lblMonth.Text + " " + Main.instance.menu.lblYear.Text;
                DateTime result;
                DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                Main.instance.menu.EmcosClear(result, DateTime.Today.AddDays(-1), RangeReferences.activeTable);
            });
            AddButtonToCommandBar("Записать данные из EMCOS", () =>
            {
                CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                string format = "dd MMMM yyyy";
                string data = RangeReferences.activeTable.ActiveDay().ToString().PadLeft(2, '0') + " " + Main.instance.menu.lblMonth.Text + " " + Main.instance.menu.lblYear.Text;
                DateTime result;
                DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                Main.instance.menu.EmcosWrite(result, result, RangeReferences.activeTable);
            });
            AddButtonToCommandBar("Очистить данные из EMCOS", () =>
            {
                CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                string format = "dd MMMM yyyy";
                string data = RangeReferences.activeTable.ActiveDay().ToString().PadLeft(2, '0') + " " + Main.instance.menu.lblMonth.Text + " " + Main.instance.menu.lblYear.Text;
                DateTime result;
                DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                Main.instance.menu.EmcosClear(result, result, RangeReferences.activeTable);
            });
            AddButtonToCommandBar("Изменить из EMCOS", Main.instance.menu.EmcosSelect);
            AddButtonToCommandBar("Удалить из EMCOS", Main.instance.menu.EmcosRemove);
            RemoveOld();
            ShowTypeMenu();
            AddButtonToCommandBar("Скрыть", Main.instance.menu.HideType);
            ChangeTypeMenuMain();
            ChangeTypeCellMenuMain();
            ChangeTypeMenuExtra();
            AddButtonToCommandBar("Ввести корректировку", () =>
            {
                using (Correct form = new Correct(RangeReferences.activeTable))
                {
                    form.ShowDialog();
                }
            }, 387);
            AddButtonToCommandBar("Добавить план", () =>
            {
                using (AddPlan form = new AddPlan(RangeReferences.activeTable))
                {
                    form.ShowDialog();
                }
            }, 213);
            AddButtonToCommandBar("Изменить код плана", () =>
            {
                using (AddPlan form = new AddPlan(RangeReferences.activeTable))
                {
                    form.ShowDialog();
                }
            }, 712);
            AddButtonToCommandBar("Удалить план", () =>
            {
                RangeReferences.activeTable.codPlan = null;
                RangeReferences.activeTable.RemovePlan();
            }, 214);
            AddButtonToCommandBar("Изменить формулу", () =>
            {
                Main.instance.menu.OpenForm();
            }, 385);
            AddButtonToCommandBar("Зависимые формулы", () =>
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
            AddButtonToCommandBar("Добавить по показаниям счетчика", () => {
                RangeReferences.activeTable.AddMeter(RangeReferences.ActiveL1);
            }, 33);
            AddButtonToCommandBar("Ввести показания счетчика", () =>
            {
                using (SCH form = new SCH(RangeReferences.activeTable))
                {
                    form.ShowDialog();
                }
            }, 205);
            AddButtonToCommandBar("Изменить коэффициент счетчика", () =>
            {
                using (ChangeCoef form = new ChangeCoef(RangeReferences.activeTable))
                {
                    form.ShowDialog();
                }
            }, 400);
            AddButtonToCommandBar("Удалить по показаниям счетчика", () => {
                RangeReferences.activeTable.RemoveMeter(RangeReferences.ActiveL1);
            });
            AddButtonToCommandBar("Сбросить", () => {
                RangeReferences.activeTable.Reset(RangeReferences.ActiveL1, RangeReferences.activeL2);
            });
            AddButtonToCommandBar("Сбросить", () => {
                RangeReferences.activeTable.ResetCell(RangeReferences.ActiveL1, RangeReferences.activeL2);
            }, tag:"Сбросить mainSubtitle");
            SpecialMenuMain();
            AddButtonToCommandBar("Удалить субъект", () => RangeReferences.activeTable.RemoveSubject(), 330);
            AddButtonToCommandBar("Удалить тип", () => RangeReferences.activeTable.RemoveChild(RangeReferences.ActiveL1));

            AddButtonToCommandBar("Добавить отступ", () =>
            {
                HeadObject ho = Main.instance.heads.HeadByRange(NewMenuBase._activeRange);
                if (ho != null)
                {
                    ho.Indent(IndentDirection.right);
                }
            }, faceid: 137);
            AddButtonToCommandBar("Удалить отступ", () =>
            {
                HeadObject ho = Main.instance.heads.HeadByRange(NewMenuBase._activeRange);
                if (ho != null)
                {
                    ho.Indent(IndentDirection.right);
                }
            }, faceid: 138);
            AddButtonToCommandBar("Удалить", () =>
            {
                HeadObject ho = Main.instance.heads.HeadByRange(NewMenuBase._activeRange);
                if (ho != null)
                {
                    if (MessageBox.Show("Это удалит всех субъектов входящих в " + ho._name + "\nВы Уверены?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        ho.Delete();
                    }
                }
            }, 1088);
        }

        static void SpecialMenuMain()
        {
            Action action;

            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Special";
            p.Tag = "Special";
            p.Visible = false;
            if (Main.instance.colors.main["subject"] == Main.instance.menu.activeColor)
            {
                AddButtonToPopUpCommandBar(ref p, "UpdateNames", RangeReferences.activeTable.UpdateNames);
            }
            AddButtonToPopUpCommandBar(ref p, "UpdateAllNames", Main.instance.references.UpdateAllNames);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllDBNames", Main.instance.references.UpdateAllDBNames);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllColors", Main.instance.references.UpdateAllColors);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllPSFormulas", Main.instance.references.UpdateAllPSFormulas);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllDBFormulas", Main.instance.references.UpdateAllDBFormulas);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllParents", Main.instance.references.UpdateAllParents, true);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllReferencesPS", Main.instance.references.UpdateAllReferencesPS);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllReferencesDB", Main.instance.references.UpdateAllReferencesDB);
            AddButtonToPopUpCommandBar(ref p, "UpdateAllLevels", Main.instance.references.UpdateAllLevels, true);
            AddButtonToPopUpCommandBar(ref p, "CheckAllRanges", Main.instance.references.CheckAllRanges);
            AddButtonToPopUpCommandBar(ref p, "ShowAllFormulas", () => {
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
                string data = "2023-02-" + Main.instance.menu.textBox1.Text;
                data = Main.instance.menu.textBox1.Text.PadLeft(2, '0') + " " + Main.instance.menu.lblMonth.Text + " " + Main.instance.menu.lblYear.Text;
                DateTime result;
                DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                if (string.IsNullOrEmpty(Main.instance.menu.textBox1.Text) || (Int32.Parse(Main.instance.menu.textBox1.Text) <= 0 && Int32.Parse(Main.instance.menu.textBox1.Text) > 31))
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
            AddButtonToPopUpCommandBar(ref p, "Test", () => {
                Main.instance.menu.EmcosWrite(new DateTime(2024, 06, 01), new DateTime(2024, 06, 19));
            });

        }

        static void AddNewL1()
        {
            CommandBarPopup p = (CommandBarPopup)CustomCellMenu.cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Добавить новый";
            p.Tag = "AddNewL1";
            p.Visible = false;
        }
        static void AddNewL2()
        {
            CommandBarPopup p = (CommandBarPopup)CustomCellMenu.cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Добавить новый";
            p.Tag = "AddNewL2";
            p.Visible = false;
        }
        static void RemoveOld()
        {
            CommandBarPopup p = (CommandBarPopup)CustomCellMenu.cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Удалить";
            p.Tag = "RemoveOld";
            p.Visible = false;
        }
        static void ShowTypeMenu()
        {
            CommandBarPopup p = (CommandBarPopup)CustomCellMenu.cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Показать";
            p.Tag = "ShowTypeMenu";
            p.Visible = false;
        }
        static void ChangeTypeMenuMain()
        {
            CommandBarPopup p = (CommandBarPopup)CustomCellMenu.cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Изменить";
            p.Tag = "ChangeTypeMenuMain";
            p.Visible = false;
        }
        static void ChangeTypeCellMenuMain()
        {
            CommandBarPopup p = (CommandBarPopup)CustomCellMenu.cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Изменить";
            p.Tag = "ChangeTypeCellMenuMain";
            p.Visible = false;
        }
        static void ChangeTypeMenuExtra()
        {
            CommandBarPopup p = (CommandBarPopup)CustomCellMenu.cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Изменить";
            p.Tag = "ChangeTypeMenuExtra";
            p.Visible = false;
        }

        #region AddButtonToCommandBar
        private static void ContextMenuClickLog(string caption)
        {
            GlobalMethods.ToLog("Нажат пункт контекстного меню '" + caption + "'");
        }
        public static void AddButtonToCommandBar(string caption, Action action, int? faceid = null, int type = 1, string tag = "", bool visible = false)
        {
            if (string.IsNullOrEmpty(tag)) tag = caption;
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(caption);
                action.Invoke();
            };
            CommandBarButton b = (CommandBarButton)cb.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Visible = visible;
            b.Tag = tag;
            if (faceid != null) b.FaceId = (int)faceid;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        public static void AddButtonToCommandBar(string caption, Action<string> action, string s1, int type = 1, string tag = "", bool visible = false)
        {
            if (string.IsNullOrEmpty(tag)) tag = caption;
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(caption);
                action.Invoke(s1);
            };
            CommandBarButton b = (CommandBarButton)cb.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Visible = visible;
            b.Tag = tag;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        public static void AddButtonToCommandBar(string caption, Action<string, string> action, string s1, string s2, int type = 1, string tag = "", bool visible = false)
        {
            if (string.IsNullOrEmpty(tag)) tag = caption;
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(caption);
                action.Invoke(s1, s2);
            };
            CommandBarButton b = (CommandBarButton)cb.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Visible = visible;
            b.Tag = tag;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }

        public static void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, Action action, int type = 1)
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
        public static void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, Action<bool> action, bool b1, int type = 1)
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
        public static void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, Action<string> action, string s1, int type = 1)
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
        public static void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, Action<string, string> action, string s1, string s2, int type = 1)
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
        public static void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, Action<string, string, bool> action, string s1, string s2, bool b1 = true, int type = 1)
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
        public static void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, Action<string, string, string?> action, string s1, string s2, string? s3 = null, int type = 1)
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
        public static void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, List<Action> action, int type = 1)
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
        public static void AddButtonToPopUpCommandBar(ref CommandBarPopup p, string caption, List<Action<string>> action, string s1, int type = 1)
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

        public static void ClearPopupMenu(CommandBarPopup p)
        {
            foreach (CommandBarControl item in p.Controls)
            {
                item.Delete();
            }
        }
    }
}

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
        public static Dictionary<string, int> CommandBarIndexes = new ();
        public static Dictionary<string, Action<CommandBarPopup>> CommandBarActions = new ();
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

        public static void RecreateCustomContextMenu()
        {
            DeleteButtonsFromMenu();
            AddButtonsToMenu();
        }

        static void DeleteButtonsFromMenu()
        {
            CommandBarIndexes.Clear();
            CommandBarActions.Clear();
            foreach (CommandBarControl item in cb.Controls)
            {
                item.Delete();
            }
        }

        static void AddButtonsToMenu()
        {
            AddButtonToCommandBar("Копировать", "Копировать", () =>
            {
                GlobalMethods.ToLog("Копирование диапазона: " + ((Excel.Range)Main.instance.xlApp.Selection).Address);
                ((Excel.Range)Main.instance.xlApp.Selection).Copy();
            }, 0019);
            AddButtonToCommandBar("Вставить", "Вставить", () =>
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
            AddButtonToCommandBar("GoTo DB", "GoTo DB", Main.instance.menu.GotoDB, 2116);
            AddButtonToCommandBar("Выделить", "Выделить", () =>
            {
                Main.instance.menu.SelectSubject();
            }, 118);
            AddButtonToCommandBar("Переместить субъект", "Переместить субъект", () =>
            {
                if (Main.instance.menu != null && Main.instance.menu.InvokeRequired)
                {
                    Main.instance.menu.Invoke(new Action(() =>
                    {
                        using (TransferSubject form = new ())
                        {
                            form.ShowDialog();
                        }
                    }));
                }
                else
                {
                    using (TransferSubject form = new ())
                    {
                        form.ShowDialog();
                    }
                }
                // using (TransferSubject form = new TransferSubject())
                // {
                //     form.ShowDialog();
                // }
            });
            AddButtonToCommandBar("Переименовать", "Переименовать", () =>
            {
                if (Main.instance.menu != null && Main.instance.menu.InvokeRequired)
                {
                    Main.instance.menu.Invoke(new Action(() =>
                    {
                        using (Rename form = new (RangeReferences.activeTable))
                        {
                            form.Show();
                        }
                    }));
                }
                else
                {
                    using (Rename form = new (RangeReferences.activeTable))
                    {
                        form.Show();
                    }
                }

                // using (Rename form = new Rename(RangeReferences.activeTable))
                // {
                //     form.ShowDialog();
                // }
            }, 7677);
            AddButtonToCommandBar("Переименовать", "Переименовать head", () =>
            {
                HeadObject ho = Main.instance.heads.HeadByRange(NewMenuBase._activeRange);
                if (Main.instance.menu != null && Main.instance.menu.InvokeRequired)
                {
                    Main.instance.menu.Invoke(new Action(() =>
                    {
                        using (Rename form = new (ho))
                        {
                            form.ShowDialog();
                        }
                    }));
                }
                else
                {
                    using (Rename form = new (ho))
                    {
                        form.ShowDialog();
                    }
                }

                // HeadObject ho = Main.instance.heads.HeadByRange(NewMenuBase._activeRange);
                // if (ho != null)
                // {
                //     using (Rename form = new Rename(ho))
                //     {
                //         form.ShowDialog();
                //     }
                // }
            }, 7677);
            
            AddButtonToCommandBar("Добавить в макетТЭП", "Добавить код для макетТЭП", () =>
            {
                ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                using (AddPlan form = new AddPlan(co))
                {
                    form.ShowDialog();
                }
            });
            AddButtonToCommandBar("Изменить код макетТЭП", "Изменить код для макетТЭП", () =>
            {
                ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                if (Main.instance.menu != null && Main.instance.menu.InvokeRequired)
                {
                    Main.instance.menu.Invoke(new Action(() =>
                    {
                        using (AddPlan form = new AddPlan(co))
                        {
                            form.ShowDialog();
                        }
                    }));
                }
                else
                {
                    using (AddPlan form = new AddPlan(co))
                    {
                        form.ShowDialog();
                    }
                }

                // ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                // using (AddPlan form = new AddPlan(co))
                // {
                //     form.ShowDialog();
                // }
            });
            AddButtonToCommandBar("Удалить из макетТЭП", "Удалить код для макетТЭП", () =>
            {
                RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].RemoveFromMTEP();
            });
            AddButtonToCommandBar("Добавить в ТЭП", "Добавить код для ТЭП", () =>
            {
                ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                if (Main.instance.menu != null && Main.instance.menu.InvokeRequired)
                {
                    Main.instance.menu.Invoke(new Action(() =>
                    {
                        using (AddTEP form = new AddTEP(co))
                        {
                            form.ShowDialog();
                        }
                    }));
                }
                else
                {
                    using (AddTEP form = new AddTEP(co))
                    {
                        form.ShowDialog();
                    }
                }

                // ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                // using (AddTEP form = new AddTEP(co))
                // {
                //     form.ShowDialog();
                // }
            });
            AddButtonToCommandBar("Изменить код ТЭП", "Изменить код для ТЭП", () =>
            {
                ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                if (Main.instance.menu != null && Main.instance.menu.InvokeRequired)
                {
                    Main.instance.menu.Invoke(new Action(() =>
                    {
                        using (AddTEP form = new AddTEP(co))
                        {
                            form.ShowDialog();
                        }
                    }));
                }
                else
                {
                    using (AddTEP form = new AddTEP(co))
                    {
                        form.ShowDialog();
                    }
                }

                // ChildObject co = RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1];
                // using (AddTEP form = new AddTEP(co))
                // {
                //     form.ShowDialog();
                // }
            });
            AddButtonToCommandBar("Удалить из ТЭП", "Удалить код для ТЭП", () =>
            {
                RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].RemoveFromTEP();
            });
            
            AddPopupToCommandBar("Добавить новый", "Добавить новый L1", (p) => 
            {
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
                if (p.accChildCount == 0) p.Visible = false;
            });
            AddPopupToCommandBar("Добавить новый", "Добавить новый L2", (p) => 
            {
                ClearPopupMenu(p);
                foreach (string n in Main.instance.colors.mainTitle.Keys)
                {
                    if (n != "по плану" && n != "по счетчику" && n != "счетчик" && n != "утвержденный" && n != "корректировка" && n != "заявка" && !RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].HasItem(n))
                    {
                        AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.DB.AddNewRange, RangeReferences.ActiveL1, n);
                    }
                }
                if (p.accChildCount == 0) p.Visible = false;
            });
            AddButtonToCommandBar("Выбрать из EMCOS", "Выбрать из EMCOS", Main.instance.menu.EmcosSelect);
            AddButtonToCommandBar("Записать из EMCOS (Все дни)", "Записать данные из EMCOS (Все дни)", () =>
            {
                CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                string format = "dd MMMM yyyy";
                string data = "01" + " " + Main.instance.menu.lblMonth.Text + " " + Main.instance.menu.lblYear.Text;
                DateTime result;
                DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                Main.instance.menu.EmcosWrite(result, DateTime.Today.AddDays(-1), RangeReferences.activeTable);
            });
            AddButtonToCommandBar("Очистить из EMCOS (Все дни)", "Очистить данные из EMCOS (Все дни)", () =>
            {
                CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                string format = "dd MMMM yyyy";
                string data = "01" + " " + Main.instance.menu.lblMonth.Text + " " + Main.instance.menu.lblYear.Text;
                DateTime result;
                DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                Main.instance.menu.EmcosClear(result, DateTime.Today.AddDays(-1), RangeReferences.activeTable);
            });
            AddButtonToCommandBar("Записать данные из EMCOS", "Записать данные из EMCOS", () =>
            {
                CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                string format = "dd MMMM yyyy";
                string data = RangeReferences.activeTable.ActiveDay().ToString().PadLeft(2, '0') + " " + Main.instance.menu.lblMonth.Text + " " + Main.instance.menu.lblYear.Text;
                DateTime result;
                DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                Main.instance.menu.EmcosWrite(result, result, RangeReferences.activeTable);
            });
            AddButtonToCommandBar("Очистить данные из EMCOS", "Очистить данные из EMCOS", () =>
            {
                CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
                string format = "dd MMMM yyyy";
                string data = RangeReferences.activeTable.ActiveDay().ToString().PadLeft(2, '0') + " " + Main.instance.menu.lblMonth.Text + " " + Main.instance.menu.lblYear.Text;
                DateTime result;
                DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
                Main.instance.menu.EmcosClear(result, result, RangeReferences.activeTable);
            });
            AddButtonToCommandBar("Изменить из EMCOS", "Изменить из EMCOS", Main.instance.menu.EmcosSelect);
            AddButtonToCommandBar("Удалить из EMCOS", "Удалить из EMCOS", Main.instance.menu.EmcosRemove);
            AddPopupToCommandBar("Удалить", "Удалить", (p) => 
            {
                ClearPopupMenu(p);
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
                if (p.accChildCount == 0) p.Visible = false;
            });
            AddPopupToCommandBar("Показать", "Показать", (p) => 
            {
                ClearPopupMenu(p);
                foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
                {
                    if (n != "код" && n != "основное" && n != "по плану" && !RangeReferences.activeTable._activeChild._activeChild.HasItem(n, SymbolType.uper))
                    {
                        AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable._activeChild.AddNewRange, RangeReferences.ActiveL1, n.ToUpper());
                    }
                }
                if (p.accChildCount == 0) p.Visible = false;
            });
            AddButtonToCommandBar("Скрыть", "Скрыть", Main.instance.menu.HideType);
            AddPopupToCommandBar("Изменить", "Изменить main", (p) => 
            {
                ClearPopupMenu(p);
                foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
                {
                    if (n != "корректировка факт" && n != "код" && n != "основное" && n != "счетчик" && !RangeReferences.activeTable._activeChild._activeChild.HasItem(n, SymbolType.lower))
                    {
                        AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.ChangeType, RangeReferences.activeL2, n);
                    }
                }
                if (p.accChildCount == 0) p.Visible = false;
            });
            AddPopupToCommandBar("Изменить", "Изменить mainSubtitle", (p) => 
            {
                ClearPopupMenu(p);
                foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
                {
                    if (n != "корректировка факт" && n != "код" && n != "основное" && n != "счетчик" && !RangeReferences.activeTable._activeChild._activeChild.HasItem(n, SymbolType.lower))
                    {
                        AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.ChangeTypeCell, RangeReferences.activeL2, n);
                    }
                }
                if (p.accChildCount == 0) p.Visible = false;
            });
            AddPopupToCommandBar("Изменить", "Изменить extra", (p) => 
            {
                ClearPopupMenu(p);
                foreach (string n in RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].childs.Keys)
                {
                    if (n != "код" && n != "основное" && n != "по плану")
                    {
                        AddButtonToPopUpCommandBar(ref p, n, RangeReferences.activeTable.ChangeType, RangeReferences.activeL2, n.ToUpper());
                    }
                }
                if (p.accChildCount == 0) p.Visible = false;
            });
            AddButtonToCommandBar("Ввести корректировку", "Ввести корректировку", () =>
            {
                if (Main.instance.menu != null && Main.instance.menu.InvokeRequired)
                {
                    Main.instance.menu.Invoke(new Action(() =>
                    {
                        using (Correct form = new Correct(RangeReferences.activeTable))
                        {
                            form.ShowDialog();
                        }
                    }));
                }
                else
                {
                    using (Correct form = new Correct(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                }

                // using (Correct form = new Correct(RangeReferences.activeTable))
                // {
                //     form.ShowDialog();
                // }
            }, 387);
            AddButtonToCommandBar("Добавить план", "Добавить план", () =>
            {
                if (Main.instance.menu != null && Main.instance.menu.InvokeRequired)
                {
                    Main.instance.menu.Invoke(new Action(() =>
                    {
                        using (AddPlan form = new AddPlan(RangeReferences.activeTable))
                        {
                            form.ShowDialog();
                        }
                    }));
                }
                else
                {
                    using (AddPlan form = new AddPlan(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                }

                // using (AddPlan form = new AddPlan(RangeReferences.activeTable))
                // {
                //     form.ShowDialog();
                // }
            }, 213);
            AddButtonToCommandBar("Изменить код плана", "Изменить код плана", () =>
            {
                if (Main.instance.menu != null && Main.instance.menu.InvokeRequired)
                {
                    Main.instance.menu.Invoke(new Action(() =>
                    {
                        using (AddPlan form = new AddPlan(RangeReferences.activeTable))
                        {
                            form.ShowDialog();
                        }
                    }));
                }
                else
                {
                    using (AddPlan form = new AddPlan(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                }

                // using (AddPlan form = new AddPlan(RangeReferences.activeTable))
                // {
                //     form.ShowDialog();
                // }
            }, 712);
            AddButtonToCommandBar("Удалить план", "Удалить план", () =>
            {
                RangeReferences.activeTable.codPlan = null;
                RangeReferences.activeTable.RemovePlan();
            }, 214);
            AddButtonToCommandBar("Изменить формулу", "Изменить формулу", () =>
            {
                Main.instance.menu.Invoke(new Action(() => 
                {
                    FormulaEditor form = new FormulaEditor(ref RangeReferences.activeTable, RangeReferences.ActiveL1);
                    form.Show();
                }));
                // Thread t = new Thread(() =>
                // {
                //     using (FormulaEditor form = new FormulaEditor(ref RangeReferences.activeTable, RangeReferences.ActiveL1))
                //     {
                //         form.FormClosed += (s, args) => 
                //         { 
                //             Application.ExitThread(); 
                //         };
                //         form.Show();
                //         Application.Run();
                //     }
                // });
                // t.SetApartmentState(ApartmentState.STA);
                // t.Start();
            }, 385);
            AddButtonToCommandBar("Зависимые формулы", "Зависимые формулы", () =>
            {
                Main.instance.menu.Invoke(new Action(() => 
                {
                    AllFormulas form = new AllFormulas(RangeReferences.activeTable, RangeReferences.ActiveL1);
                    form.Show();
                }));
            });
            AddButtonToCommandBar("Добавить по показаниям счетчика", "Добавить по показаниям счетчика", () => {
                RangeReferences.activeTable.AddMeter(RangeReferences.ActiveL1);
            }, 33);
            AddButtonToCommandBar("Ввести показания счетчика", "Ввести показания счетчика", () =>
            {
                if (Main.instance.menu != null && Main.instance.menu.InvokeRequired)
                {
                    Main.instance.menu.Invoke(new Action(() =>
                    {
                        using (SCH form = new SCH(RangeReferences.activeTable))
                        {
                            form.ShowDialog();
                        }
                    }));
                }
                else
                {
                    using (SCH form = new SCH(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                }

                // using (SCH form = new SCH(RangeReferences.activeTable))
                // {
                //     form.ShowDialog();
                // }
            }, 205);
            AddButtonToCommandBar("Изменить коэффициент счетчика", "Изменить коэффициент счетчика", () =>
            {
                if (Main.instance.menu != null && Main.instance.menu.InvokeRequired)
                {
                    Main.instance.menu.Invoke(new Action(() =>
                    {
                        using (ChangeCoef form = new ChangeCoef(RangeReferences.activeTable))
                        {
                            form.ShowDialog();
                        }
                    }));
                }
                else
                {
                    using (ChangeCoef form = new ChangeCoef(RangeReferences.activeTable))
                    {
                        form.ShowDialog();
                    }
                }

                // using (ChangeCoef form = new ChangeCoef(RangeReferences.activeTable))
                // {
                //     form.ShowDialog();
                // }
            }, 400);
            AddButtonToCommandBar("Удалить по показаниям счетчика", "Удалить по показаниям счетчика", () => {
                RangeReferences.activeTable.RemoveMeter(RangeReferences.ActiveL1);
            });
            AddButtonToCommandBar("Сбросить", "Сбросить", () => {
                RangeReferences.activeTable.Reset(RangeReferences.ActiveL1, RangeReferences.activeL2);
            });
            AddButtonToCommandBar("Сбросить", "Сбросить mainSubtitle", () => {
                RangeReferences.activeTable.ResetCell(RangeReferences.ActiveL1, RangeReferences.activeL2);
            });
            AddButtonToCommandBar("Удалить субъект", "Удалить субъект", () => RangeReferences.activeTable.RemoveSubject(), 330);
            AddButtonToCommandBar("Удалить тип", "Удалить тип", () => RangeReferences.activeTable.RemoveChild(RangeReferences.ActiveL1));
            AddButtonToCommandBar("Добавить отступ", "Добавить отступ", () => 
            {
                HeadObject ho = Main.instance.heads.HeadByRange(NewMenuBase._activeRange);
                if (ho != null)
                {
                    ho.Indent(IndentDirection.right);
                }
            }, faceid: 137);
            AddButtonToCommandBar("Удалить отступ", "Удалить отступ", () =>
            {
                HeadObject ho = Main.instance.heads.HeadByRange(NewMenuBase._activeRange);
                if (ho != null)
                {
                    ho.Indent(IndentDirection.right);
                }
            }, faceid: 138);
            AddButtonToCommandBar("Удалить", "Удалить head", () =>
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
            SpecialMenuMain();
        }
        static void SpecialMenuMain()
        {
            Action action;
            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true);
            p.Caption = "Special";
            p.Visible = false;
            CommandBarIndexes.Add("Special", p.Index);
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
                Main.instance.menu.Invoke(new Action(() => 
                {
                    AllFormulas form = new AllFormulas();
                    form.Show();
                }));
                // Thread t = new Thread(() =>
                // {
                //     using (AllFormulas form = new AllFormulas())
                //     {
                //         form.FormClosed += (s, args) =>
                //         {
                //             Application.ExitThread();
                //         };
                //         form.Show();
                //         Application.Run();
                //     }
                // });
                // t.SetApartmentState(ApartmentState.STA);
                // t.Start();
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

        #region AddButtonToCommandBar
        private static void ContextMenuClickLog(string caption)
        {
            GlobalMethods.ToLog("Нажат пункт контекстного меню '" + caption + "'");
        }
        public static void AddButtonToCommandBar(string caption, string strID, Action action, int? faceid = null, int type = 1, bool visible = false)
        {
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(caption);
                action.Invoke();
            };
            CommandBarButton b = (CommandBarButton)cb.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Visible = visible;
            CommandBarIndexes.Add(strID, b.Index);
            if (faceid != null) b.FaceId = (int)faceid;
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        public static void AddButtonToCommandBar(string caption, string strID, Action<string> action, string s1, int type = 1, bool visible = false)
        {
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(caption);
                action.Invoke(s1);
            };
            CommandBarButton b = (CommandBarButton)cb.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Visible = visible;
            CommandBarIndexes.Add(strID, b.Index);
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        public static void AddButtonToCommandBar(string caption, string strID, Action<string, string> action, string s1, string s2, int type = 1, bool visible = false)
        {
            CommandBarButtonClick newAction = (CommandBarButton commandBarButton, ref bool cancel) =>
            {
                ContextMenuClickLog(caption);
                action.Invoke(s1, s2);
            };
            CommandBarButton b = (CommandBarButton)cb.Controls.Add(Type: type, Temporary: true);
            b.Caption = caption;
            b.Visible = visible;
            CommandBarIndexes.Add(strID, b.Index);
            b.Click += new _CommandBarButtonEvents_ClickEventHandler(newAction);
        }
        public static void AddPopupToCommandBar(string caption, string strID, int type = 10, bool visible = false)
        {
            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: type, Temporary: true);
            p.Caption = caption;
            p.Visible = visible;
            CommandBarIndexes.Add(strID, p.Index);
        }
        public static void AddPopupToCommandBar(string caption, string strID, Action<CommandBarPopup> action, int type = 10, bool visible = false)
        {
            CommandBarPopup p = (CommandBarPopup)cb.Controls.Add(Type: type, Temporary: true);
            p.Caption = caption;
            p.Visible = visible;
            CommandBarIndexes.Add(strID, p.Index);
            CommandBarActions.Add(strID, action);
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

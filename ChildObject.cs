using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;

namespace Meter
{
    public class ChildObject : ReferencesParent
    {
        [JsonIgnore]
        private Excel.Range _range, _head, _body;
        [JsonIgnore]
        public Excel.Worksheet ws;
        public Level _level;
        public string? parentID, firstParentID;
        [JsonIgnore]
        int? _emcosID, _emcosMLID, _codMaketTEP, _codTEP;
        public int? emcosID
        {
            get
            {
                return _emcosID;
            }
            set
            {
                if (Main.loading == false) GlobalMethods.ToLog("Изменен emcosID для субъекта {" + GetFirstParent._name + "} " + _name + " с '" + emcosID + "' на '" + value + "'");
                _emcosID = value;
            }
        }
        public int? emcosMLID
        {
            get
            {
                return _emcosMLID;
            }
            set
            {
                if (Main.loading == false) GlobalMethods.ToLog("Изменен emcosMLID для субъекта {" + GetFirstParent._name + "} " + _name + " с '" + emcosMLID + "' на '" + value + "'");
                _emcosMLID = value;
            }
        }
        public int? codMaketTEP
        {
            get
            {
                return _codMaketTEP;
            }
            set
            {
                if (Main.loading == false) GlobalMethods.ToLog("Изменен _codMaketTEP для субъекта {" + GetFirstParent._name + "} " + _name + " с '" + _codMaketTEP + "' на '" + value + "'");
                _codMaketTEP = value;
            }
        }
        public int? codTEP
        {
            get
            {
                return _codTEP;
            }
            set
            {
                if (Main.loading == false) GlobalMethods.ToLog("Изменен _codTEP для субъекта {" + GetFirstParent._name + "} " + _name + " с '" + _codTEP + "' на '" + value + "'");
                _codTEP = value;
            }
        }
        public string WSName 
        {
            get
            {
                if (ws is null)
                    return null;
                else
                {
                    return ws.CodeName;
                }
            }
            set
            {   if (value is not null)
                {
                    if (value == "PS")
                    {
                        ws = Main.instance.wsCh;
                    }
                    else if(value == "DB")
                    {
                        ws = Main.instance.wsDb;
                    }
                }
            }
        }
        public string rangeAddress
        {
            get
            {
                if (_range is null)
                    return null;
                else
                    try
                    {
                        return _range.Address;
                    }
                    catch (Exception)
                    {

                        return null;
                    }
            }
            set
            {   if (value is not null)
                    _range = WS.Range[value];
            }
        }
        public string headAddress 
        {
            get
            {
                if (_head is null)
                    return null;
                else
                    try
                    {
                        return _head.Address;
                    }
                    catch (Exception)
                    {

                        return null;
                    }
                    
            }
            set
            {   if (value is not null)
                    _head = WS.Range[value];
            }
        }
        public string bodyAddress 
        {
            get
            {
                if (_body is null)
                    return null;
                else
                    try
                    {
                        return _body.Address;
                    }
                    catch (Exception)
                    {

                        return null;
                    }
                    
            }
            set
            {   if (value is not null)
                    _body = WS.Range[value];
            }
        }

        [JsonIgnore]
        public Excel.Range Range 
        {
            get
            {
                if (_range == null && rangeAddress != null)
                {
                    _range = WS.Range[rangeAddress];
                }
                return _range;
            } 
            set
            {
                _range = value;
                rangeAddress = value.Address;
            }
        }
        [JsonIgnore]
        public Excel.Range Head 
        {
            get
            {
                if (_head == null && headAddress != null)
                {
                    _head = WS.Range[headAddress];
                }
                return _head;
            } 
            set
            {
                _head = value;
                headAddress = value.Address;
            }
        }
        [JsonIgnore]
        public Excel.Range Body 
        {
            get
            {
                if (_body == null && bodyAddress != null)
                {
                    _body = WS.Range[bodyAddress];
                }
                return _body;
            } 
            set
            {
                _body = value;
                bodyAddress = value.Address;
            }
        }
        [JsonIgnore]
        public Excel.Worksheet WS
        {
            get
            {
                if (ws == null && WSName != null)
                {
                    if (WSName == "PS")
                    {
                        ws = Main.instance.wsCh;
                    }
                    else if(WSName == "DB")
                    {
                        ws = Main.instance.wsDb;
                    }
                }
                return ws;
            } 
            set
            {
                ws = value;
                WSName = value.CodeName;
            }
        }
        [JsonIgnore]
        public ReferenceObject GetFirstParent
        {
            get
            {
                return (ReferenceObject)RangeReferences.idDictionary[firstParentID];
            }
        }
        [JsonIgnore]
        public Excel.Range LastColumn
        {
            get
            {
                return (Excel.Range)_head.Columns[_head.Columns.Count];
            }
        }
        
        public ChildObject()
        {
            if (ID == null)
            {
                ID = Guid.NewGuid().ToString();
                while (RangeReferences.idDictionary.ContainsKey(ID))
                {
                    ID = Guid.NewGuid().ToString();
                }
            }
        }

        public T GetParent<T> () where T : ReferencesParent
        {
            return (T)RangeReferences.idDictionary[parentID];
        }
        
        public void WriteValue(int day, string val)
        {
            if (WS.CodeName == "DB")
            {
                Excel.Range r = RangeByDay(day);
                r.Value = val;
                GlobalMethods.ToLog("Запись в Базу данных ячейка " + r.Address + " Субект {" + GetFirstParent._name + "} " + GetParent<ChildObject>()._name + " " + _name + " день " + day + " значение '" + val + "'");
            }
            
        }
        public bool HasItem(string name, SymbolType symbolType = SymbolType.uperandlower)
        {
            if (childs is not null)
            {
                if (symbolType == SymbolType.lower)
                {
                    if (childs.ContainsKey(name.ToLower()))
                    {
                        return true;
                    }
                }
                else if (symbolType == SymbolType.uper)
                {
                    if (childs.ContainsKey(name.ToUpper()))
                    {
                        return true;
                    }
                }
                else if (symbolType == SymbolType.uperandlower)
                {
                    if (childs.ContainsKey(name.ToLower()) || childs.ContainsKey(name.ToUpper()))
                    {
                        return true;
                    }
                }
                
                foreach (ChildObject c in childs.Values)
                {
                    if (c.HasItem(name, symbolType))
                    {
                        return true;
                    }
                }
            }   
            return false;
        }
        public bool HasRange(Excel.Range range) 
        {
            if (WS.CodeName == "PS")
            {
                return Main.instance.xlApp.Intersect(_range, range) != null;
            }
            else
            {
                return false;
            }
        }
        public Excel.Range RangeByDay(int day)
        {
            int c;
            int k = 0;
            if (WS.CodeName == "PS") 
            {
                k = 0;
            }
            else if (ws.CodeName == "DB")
            {
                k = 1;
            }
            if (day <= 10)
            {
                c = day + k;
            }
            else if (day <= 20)
            {
                c = day + k + 1;
            }
            else
            {
                c = day + k + 3;
            }
            return (Excel.Range)_body.Cells[c, 1];
        }
        public int? DayByRange(Excel.Range range) 
        {
            int? day = null;
            if (childs == null)
            {
                int rowNumber = range.Row;
                int firstRow = ((Excel.Range)_body.Cells[1, 1]).Row;
                int d = rowNumber - firstRow + 1;
                if (d >= 1 && d <= 10)
                {
                    day = d;
                }
                else if (d >= 11 && d <= 21)
                {
                    day = d - 1;
                }
                else if (d >= 24 && d <= 34)
                {
                    day = d - 3;
                }
            }
            return day;
        }
        public void AddNewRange(string nameL1, string nameL2, bool stopall = true)
        {
            GlobalMethods.ToLog("Для субъекта " + GetFirstParent._name + " на листе " + WS.Name + " добавлен " + nameL1 + " " + nameL2);
            if (stopall == true) Main.instance.StopAll();
            int row, column, sizeColumn, sizeRow;
            string adr;

            sizeRow = _range.Rows.Count + 3;
            row = _head.Offset[-3].Row;

            ChildObject level0 = this, level1 = null;
            ChildObject l1, l2;
            if (childs.ContainsKey(nameL1))
            {
                if (childs[nameL1].childs.ContainsKey(nameL2))
                {
                    return;
                }
                else
                {
                    level1 = childs[nameL1];
                    l1 = childs[nameL1];
                    column = childs[nameL1].LastColumn.Offset[0, 1].Column;
                }
            }
            else
            {
                level1 = null;
                column = LastColumn.Offset[0, 1].Column;
            }

            if (LastColumn.Column >= column)
            {
                level0 = null;
            }
            if (level1 != null && level1.LastColumn.Column >= column)
            {
                level1 = null;
            }

            Excel.Range r = ((Excel.Range)WS.Cells[row, column]).Resize[sizeRow];
            adr = r.Resize[r.Rows.Count - 3].Offset[3].Address;
            r.Insert(Shift:Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            r = WS.Range[adr];

            if (level0 != null)
            {
                level0.Resize(1, false, stopall);
            }
            if (level1 != null)
            {
                level1.Resize(1,false, stopall);
            }
            else
            {
                l1 = new ChildObject()
                {
                    WS = this.WS,
                    _name = nameL1,
                    Range = r.Resize[r.Rows.Count - 1].Offset[1],
                    Head = r.Resize[1].Offset[1],
                    Body = r.Resize[r.Rows.Count - 2].Offset[2],
                    parentID = ID,
                    firstParentID = firstParentID,
                    childs = new Dictionary<string, ChildObject>(),
                    _level = Level.level1
                };
                
                l1._head.Value = nameL1;
                childs.Add(nameL1,l1);
                RangeReferences.idDictionary.Add(l1.ID, l1);
            }

            r = r.Resize[r.Rows.Count - 2].Offset[2];

            l2 = new ChildObject()
            {
                WS = this.WS,
                _name = nameL2,
                Range = r,
                Head = r.Resize[1],
                Body = r.Resize[r.Rows.Count - 1].Offset[1],
                parentID = childs[nameL1].ID,
                firstParentID = childs[nameL1].firstParentID,
                _level = Level.level2
            };
            l2._head.Value = nameL2;
            childs[nameL1].childs.Add(nameL2, l2);
            RangeReferences.idDictionary.Add(l2.ID, l2);

            if (WS.CodeName != "DB") 
            {
                childs[nameL1]._head.Interior.Color = Main.instance.colors.main[nameL1];
                l2.UpdateColors();
                l2.UpdateFormulas(stopall);
                try
                {
                    GetFirstParent.UpdateHeads(false);
                }
                catch
                {
                    GlobalMethods.ToLog("Err");
                }
            }
            UpdateAllBorders();
            
            if (stopall == true) Main.instance.ResumeAll();
        }
        public void CreateNewRange(string nameL1, string nameL2, bool stopall = true)
        {
            GlobalMethods.ToLog("Для субъекта " + GetFirstParent._name + " на листе " + WS.Name + " добавлен " + nameL1 + " " + nameL2);
            if (stopall == true) Main.instance.StopAll();
            int row, column, sizeColumn, sizeRow;
            string adr;

            sizeRow = _range.Rows.Count + 3;
            row = _head.Offset[-3].Row;

            ChildObject l1, l2;

            column = LastColumn.Column;
            Excel.Range r = Body;

            l1 = new ChildObject()
            {
                WS = this.WS,
                _name = nameL1,
                Range = r,
                Head = r.Resize[1],
                Body = r.Resize[r.Rows.Count - 1].Offset[1],
                parentID = ID,
                firstParentID = firstParentID,
                childs = new Dictionary<string, ChildObject>(),
                _level = Level.level1
            };

            l1._head.Value = nameL1;
            childs.Add(nameL1, l1);
            RangeReferences.idDictionary.Add(l1.ID, l1);

            r = l1.Body;

            l2 = new ChildObject()
            {
                WS = this.WS,
                _name = nameL2,
                Range = r,
                Head = r.Resize[1],
                Body = r.Resize[r.Rows.Count - 1].Offset[1],
                parentID = childs[nameL1].ID,
                firstParentID = childs[nameL1].firstParentID,
                _level = Level.level2
            };
            l2._head.Value = nameL2;
            childs[nameL1].childs.Add(nameL2, l2);
            RangeReferences.idDictionary.Add(l2.ID, l2);

            if (WS.CodeName != "DB")
            {
                childs[nameL1]._head.Interior.Color = Main.instance.colors.main[nameL1];
                l2.UpdateColors();
                l2.UpdateFormulas(stopall);
            }
            UpdateAllBorders();
            if (stopall == true) Main.instance.ResumeAll();
        }
        public void RemoveRange(string nameL1, string nameL2)
        {
            GlobalMethods.ToLog("Для субъекта " + GetFirstParent._name + " на листе " + WS.Name + " удален " + nameL1 + " " + nameL2);
            if (_level == Level.level0)
            {
                childs[nameL1].childs[nameL2].Remove();
                if (WSName == "DB")
                {
                    if (GetParent<ReferenceObject>().PS.childs[nameL1].HasItem(nameL2, SymbolType.lower))
                    {
                        GetParent<ReferenceObject>().ChangeType(nameL2, "ручное", nameL1);
                    }
                    if (GetParent<ReferenceObject>().PS.childs[nameL1].HasItem(nameL2, SymbolType.uper))
                    {
                        GetParent<ReferenceObject>().PS.childs[nameL1].childs[nameL2.ToUpper()].Remove();
                    }
                    GetParent<ReferenceObject>().PS.childs[nameL1].ChangeCod();
                }
            }
        }
        public void Remove(bool stopall = true)
        {
            string nameL1 = "", nameL2 = "", nameL0 = "";
            if (stopall) Main.instance.StopAll();
            int a = 0;
            int b = 0;
            int c;
            if (WSName == "PS")
            {
                a = 3;
            }
            if (_level == Level.level2)
            {
                b = 2;
                nameL2 = _name;
                nameL1 = GetParent<ChildObject>()._name;
                nameL0 = GetFirstParent._name;
            }
            else if (_level == Level.level1)
            {
                b = 1;
                nameL1 = _name;
                nameL0 = GetFirstParent._name;
            }
            GlobalMethods.ToLog("Для субъекта " + GetFirstParent._name + " на листе " + WS.Name + " удален " + nameL1 + " " + nameL2);
            c = a + b;
            Excel.Range r = _range.Resize[_range.Rows.Count + c].Offset[-c];
            string _id = GetFirstParent.ID;
            ClearDatas();
            r.Delete(Shift:Excel.XlDeleteShiftDirection.xlShiftToLeft); 
            if (_level != Level.level0)
            {
                GetFirstParent.PS.Head.Value = nameL0;
                GetFirstParent.PS.Head.Value = nameL0;
            }
            RangeReferences.idDictionary[_id].UpdateBorders();
            //Marshal.ReleaseComObject(r);
            if (stopall) Main.instance.ResumeAll();
        }

        public void UpdateChilds()
        {
            Head = Range.Resize[1];
            Body = Range.Resize[Range.Rows.Count - 1].Offset[1];
            Excel.Range r = ((Excel.Range)Head.Cells[1, 1]);
            _name = (string)r.Value;
            //Marshal.ReleaseComObject(r);

            int maxColumn = LastColumn.Column;
            int resizeColumns;
            int rowsCount = _range.Rows.Count;
            int bodyRowsCount = _body.Rows.Count;
            int minBodyRows = WS.CodeName == "PS" ? 37 : 38;
            if (rowsCount > minBodyRows)
            {
                if (childs != null) 
                    childs.Clear();
                else
                    childs = new Dictionary<string, ChildObject>();

                Excel.Range rng1 = ((Excel.Range)_body.Cells[1, 1]);
                Excel.Range rng2 = null;
                
                while (rng1.Column <= maxColumn)
                {
                    resizeColumns = rng1.MergeArea.Columns.Count;
                    rng2 = rng1.Resize[bodyRowsCount, resizeColumns];
                    string name = (string)((Excel.Range)rng2.Cells[1, 1]).Value;
                    string address = rng2.Address;
                    childs.Add(name, new ChildObject()
                    {
                        WS = this.WS,
                        Range = rng2,
                        parentID = ID,
                        firstParentID = firstParentID,
                    });
                    //Connector.tasks.Add(Task.Run(() => childs[name].Generate()));
                    rng1 = rng1.Offset[0, 1];
                }
                //if (rng1 != null) Marshal.ReleaseComObject(rng1);
                //if (rng2 != null) Marshal.ReleaseComObject(rng2);
            }
            if (RangeReferences.idDictionary.ContainsKey(ID))
            {
                if (((ChildObject)RangeReferences.idDictionary[ID]).Range.Address != Range.Address)
                {
                    RangeReferences.IDErrors1.Add((ChildObject)RangeReferences.idDictionary[ID]);
                    RangeReferences.IDErrors2.Add(this);
                }
            }
            else
            {
                RangeReferences.idDictionary.Add(ID, this);
            }
            
            if (childs != null)
            {
                foreach (ChildObject c in childs.Values)
                {
                    c.UpdateChilds();
                }
            }
        }
        public void ChangeColorCell()
        {
            if (_level != Level.level2) return;
        }
        public void UpdateFormulaCell()
        {
            if (_level != Level.level2) return;
        }
        public void UpdateColors()
        {
            if (WS.CodeName == "DB") return;
            Color? cT = null, cS = null, s = Main.instance.colors.sumColor;
            
            if (childs == null)
            {
                if (Main.instance.colors.mainTitle.ContainsKey(_name))
                {
                    cT = Main.instance.colors.mainTitle[_name];
                    cS = Main.instance.colors.mainSubtitle[_name];
                }
                else if (Main.instance.colors.extraTitle.ContainsKey(_name))
                {
                    cT = Main.instance.colors.extraTitle[_name];
                    cS = Main.instance.colors.extraSubtitle[_name];
                }
                if (cT != null && cS != null)
                {
                    _head.Interior.Color = cT;
                    _body.Interior.Color = cS;

                    ((Excel.Range)Body.Rows[11]).Interior.Color = s;
                    ((Excel.Range)Body.Rows[22]).Interior.Color = s;
                    ((Excel.Range)Body.Rows[23]).Interior.Color = s;
                    ((Excel.Range)Body.Rows[35]).Interior.Color = s;
                    ((Excel.Range)Body.Rows[36]).Interior.Color = s;

                    //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[11]).Interior.Color);
                    //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[22]).Interior.Color);
                    //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[23]).Interior.Color);
                    //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[35]).Interior.Color);
                    //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[36]).Interior.Color);
                }
            }
            else
            {
                if (_level == Level.level0)
                {
                    _head.Interior.Color = Main.instance.colors.main["subject"];
                }
                else
                {
                    if (Main.instance.colors.main.ContainsKey(_name))
                        _head.Interior.Color = Main.instance.colors.main[_name];
                }
            }
        }
        public void UpdateAllColors()
        {

            UpdateColors();

            if (childs != null)
            {
                foreach (ChildObject c in childs.Values)
                {
                    c.UpdateAllColors();
                }
            }
        }
        public override void UpdateBorders()
        {
            if (_level == Level.level0)
            {
                Body.Borders.LineStyle = XlLineStyle.xlContinuous;
                
            }
            Body.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            //Body.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            Head.Borders.Weight = 3;

            if (_level == Level.level2)
            {
                if (WS.CodeName == "PS")
                {
                    ((Excel.Range)Body.Rows[11]).BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    ((Excel.Range)Body.Rows[22]).BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    ((Excel.Range)Body.Rows[23]).BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    ((Excel.Range)Body.Rows[35]).BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    ((Excel.Range)Body.Rows[36]).BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                }
                else if (WS.CodeName == "DB")
                {
                    ((Excel.Range)Body.Rows[1]).BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    ((Excel.Range)Body.Rows[12]).BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    ((Excel.Range)Body.Rows[23]).BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    ((Excel.Range)Body.Rows[24]).BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium); 
                    ((Excel.Range)Body.Rows[36]).BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium); 
                    ((Excel.Range)Body.Rows[37]).BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                }
            }
        }
        public void UpdateAllBorders()
        {
            UpdateBorders();

            if (childs != null)
            {
                foreach (ChildObject c in childs.Values)
                {
                    c.UpdateAllBorders();
                }
            }
        }
        public void UpdateFormulas(bool stopall = true, bool clear = true)
        {
            if (childs == null && WS.CodeName != "DB")
            {
                //Main.instance.stopped = true;
                if (stopall == true) Main.instance.StopAll();
                if (Main.instance.colors.mainTitle.ContainsKey(_name))
                {
                    string ad = GetFirstParent.DB.WS.Name + "!" + ((Excel.Range)GetFirstParent.DB.childs[GetParent<ChildObject>()._name].childs["основное"]._body.Cells[2, 1]).Address[false, false];
                    
                    Body.Formula = "=IF(OR(" + ad + @"="""",ISERROR(" + ad + @")),""""," + ad + ")";
                    // Body.Formula = "=" + GetFirstParent.DB.WS.Name + "!" + ((Excel.Range)GetFirstParent.DB.childs[GetParent<ChildObject>()._name].childs["основное"]._body.Cells[2, 1]).Address[false, false];
                }
                else if (Main.instance.colors.extraTitle.ContainsKey(_name))
                {
                    string ad = GetFirstParent.DB.WS.Name + "!" + ((Excel.Range)GetFirstParent.DB.childs[GetParent<ChildObject>()._name].childs[_name.ToLower()]._body.Cells[2, 1]).Address[false, false];
                    
                    Body.Formula = "=IF(OR(" + ad + @"="""",ISERROR(" + ad + @")),""""," + ad + ")";
                    // Body.Formula = "=" + GetFirstParent.DB.WS.Name + "!" + ((Excel.Range)GetFirstParent.DB.childs[GetParent<ChildObject>()._name].childs[_name.ToLower()]._body.Cells[2, 1]).Address[false, false];
                }
                ((Excel.Range)Body.Rows[11]).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)";
                ((Excel.Range)Body.Rows[22]).FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)";
                ((Excel.Range)Body.Rows[23]).FormulaR1C1 = "=R[-12]C+R[-1]C";
                ((Excel.Range)Body.Rows[35]).FormulaR1C1 = "=SUM(R[-11]C:R[-1]C)";
                ((Excel.Range)Body.Rows[36]).FormulaR1C1 = "=R[-13]C+R[-1]C";
                //Main.instance.stopped = false;
                if (stopall == true) Main.instance.ResumeAll();
            }
            else if (childs == null && WS.CodeName == "DB")
            {
                if (_name == "счетчик")
                {
                    if (clear == true)
                    {
                        object val = ((Excel.Range)Body.Cells[Body.Cells.Count]).Value;
                        Body.ClearContents();

                        ((Excel.Range)Body.Rows[1]).Value = val;
                    }
                    ((Excel.Range)Body.Rows[12]).FormulaR1C1 = "=R[-1]C";
                    ((Excel.Range)Body.Rows[23]).FormulaR1C1 = "=R[-1]C";
                    ((Excel.Range)Body.Rows[24]).FormulaR1C1 = "=R[-1]C";
                    ((Excel.Range)Body.Rows[36]).FormulaR1C1 = "=R[-1]C";
                    ((Excel.Range)Body.Rows[37]).FormulaR1C1 = "=R[-1]C";

                    //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[12]).FormulaR1C1);
                    //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[23]).FormulaR1C1);
                    //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[24]).FormulaR1C1);
                    //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[36]).FormulaR1C1);
                    //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[37]).FormulaR1C1);
                }
                if (_name == "формула")
                {
                    return;
                    // object val = ((Excel.Range)Body.Cells[Body.Cells.Count]).Value;
                    // Body.ClearContents();
                }
                else if (_name == "по счетчику")
                {
                    string? coef = GetFirstParent.meterCoef;
                    string address = ((Excel.Range)Body.Cells[1, 1]).Address[ReferenceStyle: XlReferenceStyle.xlR1C1];

                    Body.FormulaR1C1 = "=IF(OR(ISBLANK(RC[-1]),ISBLANK(" + address + @")),"""",(RC[-1]-R[-1]C[-1])*" + address + ")";
                    Body.Replace("$", "", LookAt: XlLookAt.xlPart);
                    
                    ((Excel.Range)Body.Rows[1]).Value = coef != null ? coef : "";
                    ((Excel.Range)Body.Rows[12]).ClearContents();
                    ((Excel.Range)Body.Rows[23]).ClearContents();
                    ((Excel.Range)Body.Rows[24]).ClearContents(); 
                    ((Excel.Range)Body.Rows[36]).ClearContents(); 
                    ((Excel.Range)Body.Rows[37]).ClearContents();

                    //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[1]).Value);
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[12]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[23]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[24]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[36]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[37]));
                }
                else if (_name == "код")
                {
                    Body.Value = "=" + ((Excel.Range)Body.Cells[1, 1]).Address;
                    Body.Replace("$", "", LookAt: XlLookAt.xlPart);

                    string nameL1, nameL2;
                    nameL1 = GetParent<ChildObject>()._name;
                    if (GetFirstParent.PS.childs.ContainsKey(nameL1))
                    {
                        GetFirstParent.PS.childs[nameL1].childs.Values.Where(n => Main.instance.colors.mainTitle.ContainsKey(n._name) == true).FirstOrDefault().ChangeCod();
                    }
                    else
                    {
                        ((Excel.Range)Body.Rows[1]).Value = 1;
                    }
                    
                    ((Excel.Range)Body.Rows[12]).ClearContents();
                    ((Excel.Range)Body.Rows[23]).ClearContents();
                    ((Excel.Range)Body.Rows[24]).ClearContents(); 
                    ((Excel.Range)Body.Rows[36]).ClearContents(); 
                    ((Excel.Range)Body.Rows[37]).ClearContents();

                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[12]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[23]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[24]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[36]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[37]));
                }
                else if (_name == "основное")
                {
                    if (GetParent<ChildObject>()._name == "план")
                    {
                        Body.FormulaR1C1 = "=@INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) + RC[-1],4,1),TRUE)";
                        // Body.FormulaR1C1 = "=IF(OR(ISERROR(INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) + RC[-1],4,1),TRUE) + RC[1]),ISBLANK(INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) + RC[-1],4,1),TRUE))),"""",INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) + RC[-1],4,1),TRUE) + RC[1])"
                    }
                    else
                    {
                        // Body.FormulaR1C1 = "=@INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) + RC[-1],4,1),TRUE) + RC[1]";
                        Body.FormulaR1C1 = @"=IF(OR(ISERROR(INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) + RC[-1],4,1),TRUE) + RC[1]),ISBLANK(INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) + RC[-1],4,1),TRUE))),"""",INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) + RC[-1],4,1),TRUE) + RC[1])";
                    }
                    ((Excel.Range)Body.Rows[1]).ClearContents();;
                    ((Excel.Range)Body.Rows[12]).ClearContents();
                    ((Excel.Range)Body.Rows[23]).ClearContents();
                    ((Excel.Range)Body.Rows[24]).ClearContents(); 
                    ((Excel.Range)Body.Rows[36]).ClearContents(); 
                    ((Excel.Range)Body.Rows[37]).ClearContents();

                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[1]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[12]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[23]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[24]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[36]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[37]));
                }
                else if (_name == "по плану")
                {
                    Body.Formula = "=" + ((Excel.Range)GetFirstParent.DB.childs["план"].childs["основное"].Body.Cells[1,1]).Address[false, false];
                    ((Excel.Range)Body.Rows[1]).ClearContents();;
                    ((Excel.Range)Body.Rows[12]).ClearContents();
                    ((Excel.Range)Body.Rows[23]).ClearContents();
                    ((Excel.Range)Body.Rows[24]).ClearContents(); 
                    ((Excel.Range)Body.Rows[36]).ClearContents(); 
                    ((Excel.Range)Body.Rows[37]).ClearContents();

                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[1]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[12]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[23]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[24]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[36]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[37]));
                }
                else if (_name == "заявка")
                {
                    // Body.FormulaR1C1 = "=SUM(RC[1],RC[2])";
                    Body.FormulaR1C1 = @"=IF(AND(ISBLANK(RC[1]),ISBLANK(RC[2])),"""",SUM(RC[1],RC[2]))";
                    ((Excel.Range)Body.Rows[1]).ClearContents();;
                    ((Excel.Range)Body.Rows[12]).ClearContents();
                    ((Excel.Range)Body.Rows[23]).ClearContents();
                    ((Excel.Range)Body.Rows[24]).ClearContents(); 
                    ((Excel.Range)Body.Rows[36]).ClearContents(); 
                    ((Excel.Range)Body.Rows[37]).ClearContents();

                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[1]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[12]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[23]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[24]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[36]));
                    Marshal.ReleaseComObject(((Excel.Range)Body.Rows[37]));
                }
                else
                {
                    if (clear == true)
                    {
                        ((Excel.Range)Body.Rows[12]).ClearContents(); 
                        ((Excel.Range)Body.Rows[23]).ClearContents();
                        ((Excel.Range)Body.Rows[24]).ClearContents();
                        ((Excel.Range)Body.Rows[36]).ClearContents();
                        ((Excel.Range)Body.Rows[37]).ClearContents();
    
                        Marshal.ReleaseComObject(((Excel.Range)Body.Rows[12]));
                        Marshal.ReleaseComObject(((Excel.Range)Body.Rows[23]));
                        Marshal.ReleaseComObject(((Excel.Range)Body.Rows[24]));
                        Marshal.ReleaseComObject(((Excel.Range)Body.Rows[36]));
                        Marshal.ReleaseComObject(((Excel.Range)Body.Rows[37]));
                    }
                }
            }
            else
            {
                foreach (ChildObject item in childs.Values)
                {
                    item.UpdateFormulas(stopall, clear);
                }
            }
        }
        public void UpdateLevels()
        {
            if (childs != null)
            {
                foreach (ChildObject item in childs.Values)
                {
                    item._level = Level.level1;
                    item.UpdateLevels1();
                }
            }
        }
        public void UpdateLevels1()
        {
            if (childs != null)
            {
                foreach (ChildObject item in childs.Values)
                {
                    item._level = Level.level2;
                }
            }
        }
        public void Resize(int column, bool newColumn = true, bool stopall = true)
        {
            if (newColumn)
            {
                Range = Range.Resize[Range.Rows.Count ,column];
                Head = Head.Resize[Head.Rows.Count ,column];
                Body = Body.Resize[Body.Rows.Count ,column];
            }
            else
            {
                Range = Range.Resize[Range.Rows.Count ,Range.Columns.Count + column];
                Head = Head.Resize[Head.Rows.Count ,Head.Columns.Count + column];
                Body = Body.Resize[Body.Rows.Count ,Body.Columns.Count + column];
            }

            if (stopall == true) Main.instance.xlApp.DisplayAlerts = false;
            Head.Merge();
            if (stopall == true) Main.instance.xlApp.DisplayAlerts = true;
        }
        public void ClearDatas()
        {
            if (childs != null)
            {
                foreach (ChildObject item in childs.Values)
                {
                    item.ClearDatas();
                }
            }
            RangeReferences.idDictionary.Remove(ID);
            GetParent<ReferencesParent>().childs.Remove(_name);
            if (Main.instance.formulas.formulas.ContainsKey(ID))
            {
                Main.instance.formulas.formulas.Remove(ID);
            }
            List<ForTags> temp = new List<ForTags>();
            // temp = Main.instance.formulas.formulas.Where(kv => (kv.Value != null && (kv.Value.Any(n => (n != null && n.ID != null && n.ID == ID))))).SelectMany(kv => kv.Value).Where(c => (c != null && c.ID != null && c.ID == ID)).ToList();
            temp = Main.instance.formulas.formulas.Values.SelectMany(f => f).Where(t => t.ID == ID).ToList();
            if (temp.Count != 0)
            {
                foreach (ForTags f in temp)
                {
                    // f.text = "#ссылка";
                    f.text += " (удален)";
                    f.ID = null;
                }
            }
            childs = null;
            _range = null;
            _body = null;
            _head = null;
            ws = null;
            ID = null;
            parentID = null;

            //Marshal.ReleaseComObject(_range);
            //Marshal.ReleaseComObject(_body);
            //Marshal.ReleaseComObject(_head);
            //Marshal.ReleaseComObject(ws);
        }
        public void UpdateNames()
        {
            Excel.Range r = ((Excel.Range)_head.Cells[1, 1]);
            _name = (string)r.Value;
            if (childs != null)
            {
                List<string> names = childs.Keys.ToList();
                foreach (string k in names)
                {
                    ChildObject c = childs[k];
                    childs[k].UpdateNames();
                    string newName = c._name;
                    childs.Remove(k);
                    childs.Add(newName, c);
                }
            }
        }
        public void UpdateDBNames()
        {
            ((Excel.Range)Head.Cells[1,1]).Value = _name;
        }
        public void UpdateReferences()
        {
            if (childs != null)
            {
                foreach (ChildObject item in childs.Values)
                {
                    item.ClearDatas();
                }
            }
            Excel.Range r = WS.UsedRange.Find(What: RangeReferences.idDictionary[firstParentID]._name, LookAt: Excel.XlLookAt.xlWhole);
            if (r != null)
            {
                int size;
                if (WS.CodeName == "PS")
                {
                    size = 39;
                }
                else
                {
                    size = 40;
                }
                r = r.MergeArea.Resize[size];

                Range = r;

                UpdateChilds();
            }
            //Marshal.ReleaseComObject(r);
        }
        public override void UpdateParents()
        {
            base.UpdateParents();
            // if (ID == null || RangeReferences.idDictionary.ContainsKey(ID))
            // {
            //     ID = Guid.NewGuid().ToString();
            //     while (RangeReferences.idDictionary.ContainsKey(ID))
            //     {
            //         ID = Guid.NewGuid().ToString();
            //     }
            // }
            RangeReferences.idDictionary.Add(ID, this);
            if (_level != Level.level2)
            {
                foreach (ChildObject item in childs.Values)
                {
                    item.parentID = ID;
                    item.firstParentID = firstParentID;
                    item.UpdateParents();
                }
            }
        }
        
        public void ChangeCod()
        {
            if (childs == null)
            {
                if (Main.instance.colors.mainTitle.ContainsKey(_name))
                {
                    int cod = ((ReferenceObject)RangeReferences.idDictionary[firstParentID]).GetCode(_name, GetParent<ChildObject>()._name);
                    Excel.Range r = ((Excel.Range)((ReferenceObject)RangeReferences.idDictionary[firstParentID]).DB.childs[GetParent<ChildObject>()._name].childs["код"].Body.Cells[1, 1]);
                    r.Value = cod;
                    //Marshal.ReleaseComObject(r);
                }
            }
            else
            {
                foreach (ChildObject item in childs.Values)
                {
                    item.ChangeCod();
                }
            }
        }

        public void ResetCode()
        {
            if (childs == null)
            {
                Excel.Range r = ((Excel.Range)((ReferenceObject)RangeReferences.idDictionary[firstParentID]).DB.childs[GetParent<ChildObject>()._name].childs["код"].Body.Cells[1, 1]);
                ((ReferenceObject)RangeReferences.idDictionary[firstParentID]).DB.childs[GetParent<ChildObject>()._name].childs["код"].Body.Value = "=" + r.Address;
                ((ReferenceObject)RangeReferences.idDictionary[firstParentID]).DB.childs[GetParent<ChildObject>()._name].childs["код"].Body.Replace("$", "", LookAt: Excel.XlLookAt.xlPart);

                Marshal.ReleaseComObject(r);

                int cod = ((ReferenceObject)RangeReferences.idDictionary[firstParentID]).GetCode(_name, GetParent<ChildObject>()._name);
                r = ((Excel.Range)((ReferenceObject)RangeReferences.idDictionary[firstParentID]).DB.childs[GetParent<ChildObject>()._name].childs["код"].Body.Cells[1, 1]);
                r.Value = cod;

                Marshal.ReleaseComObject(r);
            }
            else
            {
                foreach (ChildObject item in childs.Values)
                {
                    item.ResetCode();
                }
            }
        }

        public void ResetCodeCell(int day)
        {
            Excel.Range r = ((Excel.Range)((ReferenceObject)RangeReferences.idDictionary[firstParentID]).DB.childs[GetParent<ChildObject>()._name].childs["код"].Body.Cells[1, 1]);
            ((ReferenceObject)RangeReferences.idDictionary[firstParentID]).DB.childs[GetParent<ChildObject>()._name].childs["код"].RangeByDay(day).Value = "=" + r.Address;
            ((ReferenceObject)RangeReferences.idDictionary[firstParentID]).DB.childs[GetParent<ChildObject>()._name].childs["код"].RangeByDay(day).Replace("$", "", LookAt: Excel.XlLookAt.xlPart);
            Marshal.ReleaseComObject(r);
        }
        public void Clear()
        {
            if (_name == "счетчик")
            {
                UpdateFormulas();
            }
            else if (_name == "по счетчику")
            {
                UpdateFormulas();
            }
            else if (_name == "формула")
            {
                UpdateFormulas();
                //Body.FormulaR1C1 = "";
                
                //((Excel.Range)Body.Rows[1]).ClearContents();;
                //((Excel.Range)Body.Rows[12]).ClearContents();
                //((Excel.Range)Body.Rows[23]).ClearContents();
                //((Excel.Range)Body.Rows[24]).ClearContents(); 
                //((Excel.Range)Body.Rows[36]).ClearContents(); 
                //((Excel.Range)Body.Rows[37]).ClearContents();

                //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[1]));
                //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[12]));
                //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[23]));
                //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[24]));
                //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[36]));
                //Marshal.ReleaseComObject(((Excel.Range)Body.Rows[37]));
            }
            else if (_name == "по плану")
            {
                UpdateFormulas();
            }
            else if (_name == "заявка")
            {
                UpdateFormulas();
            }
            else if (_name == "код")
            {
                UpdateFormulas();
            }
            else if (_name == "основное")
            {
                UpdateFormulas();
            }
            else
            {
                Body.ClearContents();
                //((Excel.Range)Body.Rows[12]).ClearContents(); 
                //((Excel.Range)Body.Rows[23]).ClearContents();
                //((Excel.Range)Body.Rows[24]).ClearContents();
                //((Excel.Range)Body.Rows[36]).ClearContents();
                //((Excel.Range)Body.Rows[37]).ClearContents();
            }
        }
        public void ClearAll()
        {
            if (_level == Level.level2)
            {
                Clear();
            }
            else
            {
                foreach (ChildObject c in childs.Values)
                {
                    c.ClearAll();
                }
            }
        }

        public void Check()
        {
            string? errVal = null, name = null;

            if (_level == Level.level0)
            {
                name = "{" + WS.CodeName + ": " + _name + "}";
            }
            else if (_level == Level.level1)
            {
                name = "{" + WS.CodeName + ": " + GetParent<ChildObject>()._name + " " + _name + "}";
            }
            else if (_level == Level.level2)
            {
                name = "{" + WS.CodeName + ": " + GetParent<ChildObject>().GetParent<ChildObject>()._name + " " + GetParent<ChildObject>()._name + " " + _name + "}";
            }

            if (_name != (string)((Excel.Range)Head.Cells[1,1]).Value)
            {
                errVal = "HeadErr";
            }
            else if (Head.Column != Body.Column || Head.Column != Range.Column)
            {
                errVal = errVal == null ? "ColumnErr: " : errVal + " ColumnErr: ";
            }
            if (errVal != null)
            {
                Main.instance.references.Errors.Add(errVal + name);
            }
            if (_level != Level.level2)
            {
                foreach (ChildObject item in childs.Values)
                {
                    item.Check();
                }
            }
            //if (_level == Level.level0)
            //{
            //    if (WS.CodeName == "PS")
            //    {
            //        ReferenceObject co = Main.instance.references.references.Values.Where(c => c != GetFirstParent && c.HasRangePS(Range) == true).FirstOrDefault();

            //        if (co != null)
            //        {
            //            Main.instance.references.Errors.Add("Дубликат: " + name + " and " + co._name);
            //        }
            //    }
            //    else if (WS.CodeName == "DB")
            //    {
            //        ReferenceObject co = Main.instance.references.references.Values.Where(c => c != GetFirstParent && c.HasRangeDB(Range) == true).FirstOrDefault();

            //        if (co != null)
            //        {
            //            Main.instance.references.Errors.Add("Дубликат: " + name + " and " + co._name);
            //        }
            //    }
            //}
        }
        public void WriteFormula(string formula)
        {
            if (_level == Level.level2)
            {
                if (formula != "")
                {
                    Body.Formula = "=IFERROR(" + formula + @","""")";
                }
                else
                {
                    Body.Formula = "";
                }

                ((Excel.Range)Body.Rows[1]).ClearContents();
                ((Excel.Range)Body.Rows[12]).ClearContents();
                ((Excel.Range)Body.Rows[23]).ClearContents();
                ((Excel.Range)Body.Rows[24]).ClearContents();
                ((Excel.Range)Body.Rows[36]).ClearContents();
                ((Excel.Range)Body.Rows[37]).ClearContents();

                Marshal.ReleaseComObject((Excel.Range)Body.Rows[1]);
                Marshal.ReleaseComObject((Excel.Range)Body.Rows[12]);
                Marshal.ReleaseComObject((Excel.Range)Body.Rows[23]);
                Marshal.ReleaseComObject((Excel.Range)Body.Rows[24]);
                Marshal.ReleaseComObject((Excel.Range)Body.Rows[36]);
                Marshal.ReleaseComObject((Excel.Range)Body.Rows[37]);
            }
            else if (_level== Level.level1)
            {
                childs["формула"].WriteFormula(formula);
            }
        }
        public void ReleaseAllComObjects()
        {
            
            if (childs != null)
            {
                foreach (ChildObject co in childs.Values)
                {
                    co.ReleaseAllComObjects();
                }
            }
            GlobalMethods.ReleseObject(Range);
            GlobalMethods.ReleseObject(Body);
            GlobalMethods.ReleseObject(Head);
            GlobalMethods.ReleseObject(WS);
            GlobalMethods.ReleseObject(_activeChild);
        }
        public void WriteToMaketTEP(int day, int? cod = null)
        {
            if (cod == null)
                cod = codMaketTEP;
            if (_level == Level.level2)
            {
                var val = RangeByDay(day).Value;
                GlobalMethods.ToLog("Записано значение " + val + " в макетТЭП для " + GetFirstParent._name + " " + GetParent<ChildObject>()._name + " по коду " + cod);
                Main.instance.wsMTEP.Range["A:A"].Find(What: cod, LookAt: XlLookAt.xlWhole).Offset[0, 1].Value = val;
            }
            else if (_level == Level.level1)
            {
                GetFirstParent.DB.childs[_name].childs["основное"].WriteToMaketTEP(day, codMaketTEP);
            }
        }
        public void WriteToTEP(int day, int row, int? cod = null, bool rewrite = false)
        {
            if (cod == null)
                cod = codTEP;
            if (_level == Level.level2)
            {
                var val = RangeByDay(day).Value;
                int column;
                GlobalMethods.ToLog("Записано значение " + val + " в ТЭП для " + GetFirstParent._name + " " + GetParent<ChildObject>()._name + " по коду " + cod);
                if (GetParent<ChildObject>()._name != "план") 
                    column = Main.instance.wsTEPm.Range["5:5"].Find(What: cod, LookAt: XlLookAt.xlWhole).Offset[0, 1].Column;
                else
                    column = Main.instance.wsTEPm.Range["5:5"].Find(What: cod, LookAt: XlLookAt.xlWhole).Column;

                ((Excel.Range)Main.instance.wsTEPm.Cells[row, column]).Value = val;

                if (rewrite == false)
                {
                    ((Excel.Range)Main.instance.wsTEPn.Cells[row, column]).Value = "=INDIRECT(ADDRESS(ROW(RC)+1,COLUMN(RC),4,1))+INDIRECT(CONCATENATE(\"" + Main.instance.wsTEPm.Name + "!\",ADDRESS(ROW(RC),COLUMN(RC),4,1)))";
                }
            }
            else if (_level == Level.level1)
            {
                GetFirstParent.DB.childs[_name].childs["основное"].WriteToTEP(day, row, codTEP, rewrite);
                if (GetFirstParent.DB.childs.ContainsKey("план"))
                    GetFirstParent.DB.childs["план"].childs["заявка"].WriteToTEP(day, row, codTEP, rewrite);
            }
        }
    }
}
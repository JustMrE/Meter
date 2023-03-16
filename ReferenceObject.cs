using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;
using Newtonsoft.Json;
using System.Threading;
using System;
using System.Globalization;
using System.Runtime.InteropServices;
using Meter.Forms;

namespace Meter
{
    public class ReferenceObject : ReferencesParent
    {
        [JsonIgnore]
        int? _codPlan;
        public int? codPlan
        {
            get
            {
                return _codPlan;
            }
            set
            {
                if (Main.loading == false) GlobalMethods.ToLog("Для субъекта {" + _name + "} изменен код плана с '" + _codPlan + "' на '" + value.ToString() + "'");
                _codPlan = value; 
            }
        }
        public string? meterCoef
        {
            get;
            set;
        }

        [JsonIgnore]
        public HeadObject HeadL0
        {
            get
            {
                string nameL0 = ((Excel.Range)((Excel.Range)PS.Head.Cells[1, 1]).Offset[-3].MergeArea.Cells[1,1]).Value as string;
                if (Main.instance.heads.heads.ContainsKey(nameL0))
                {
                    return Main.instance.heads.heads[nameL0];
                }
                else
                {
                    return null;
                }
            }
        }
        [JsonIgnore]
        public HeadObject HeadL1
        {
            get
            {
                string nameL1 = ((Excel.Range)((Excel.Range)PS.Head.Cells[1, 1]).Offset[-2].MergeArea.Cells[1,1]).Value as string;
                HeadObject ho = HeadL0;
                if (ho != null && ho.childs.ContainsKey(nameL1))
                {
                    return ho.childs[nameL1];
                }
                else
                {
                    return null;
                }
            }
        }
        [JsonIgnore]
        public HeadObject HeadL2
        {
            get
            {
                string nameL2 = ((Excel.Range)((Excel.Range)PS.Head.Cells[1, 1]).Offset[-1].MergeArea.Cells[1,1]).Value as string;
                HeadObject ho = HeadL1;
                if (ho != null && ho.childs.ContainsKey(nameL2))
                {
                    return ho.childs[nameL2];
                }
                else
                {
                    return null;
                }
            }
        }
        [JsonIgnore]
        public ChildObject DB
        {
            get
            {
                return childs["DB"];
            }
        }
        [JsonIgnore]
        public ChildObject PS
        {
            get
            {
                return childs["PS"];
            }
        }
        [JsonIgnore]
        public bool HasEmcosID
        {
            get
            {
                return DB.childs.Values.Any(n => n.emcosID != null);
            }
        }

        public ReferenceObject()
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
        public ReferenceObject(string name, string nameL1, string address, bool insert = true, bool stopall = true) : this()
        {
            if (stopall == true) Main.instance.StopAll();

            RangeReferences.idDictionary.Add(ID, this);
            _name = name;
            childs = new Dictionary<string, ChildObject>();

            Excel.Range range;
            string adr, adrPS, adrDB;

            adr = Main.instance.wsDb.Range["B2"].Value as string;
            range = Main.instance.wsDb.Range[adr];
            range = range.Offset[-2].Resize[40];
            adr = range.Address[false, false];
            range.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            range = Main.instance.wsDb.Range[adr];
            ChildObject db = new ChildObject()
            {
                _name = name,
                WS = Main.instance.wsDb,
                rangeAddress = range.Address[false, false],
                headAddress = range.Resize[1].Address[false, false],
                bodyAddress = range.Resize[range.Rows.Count - 1].Offset[1].Address[false, false],
                _level = Level.level0,
                parentID = ID,
                firstParentID = ID,
                childs = new Dictionary<string, ChildObject>(),
            };
            db.Head.Value = name;
            RangeReferences.idDictionary.Add(db.ID, db);

            adr = address;
            range = Main.instance.wsCh.Range[adr];
            range = range.Offset[-3].Resize[42];
            adr = range.Offset[3].Resize[39].Address[false, false];
            if (insert) range.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            range = Main.instance.wsCh.Range[adr];

            ChildObject ps = new ChildObject()
            {
                _name = name,
                WS = Main.instance.wsCh,
                rangeAddress = range.Address[false, false],
                headAddress = range.Resize[1].Address[false, false],
                bodyAddress = range.Resize[range.Rows.Count - 1].Offset[1].Address[false, false],
                _level = Level.level0,
                parentID = ID,
                firstParentID = ID,
                childs = new Dictionary<string, ChildObject>(),
            };
            ps.Head.Value = name;
            RangeReferences.idDictionary.Add(ps.ID, ps);
            
            childs.Add("PS", ps);
            childs.Add("DB", db);

            CreateNewDBL1StandartOther(nameL1, stopall);
            CreateNewPS(nameL1, "ручное", stopall);

             if (stopall == true) Main.instance.ResumeAll();
        }

        public int? ActiveDay()
        {
            return _activeChild._activeChild._activeChild.DayByRange(NewMenuBase._activeRange);
        }
        public bool HasRange(Excel.Range rng)
        {
            return Main.instance.xlApp.Intersect(PS.Range, rng) != null;
        }
        public void WriteToDB(string nameL1, string nameL2, int day, string val)
        {
            if (!string.IsNullOrEmpty(val))
            {
                double doubleVal;
                CultureInfo culture = CultureInfo.InvariantCulture;
                if (double.TryParse(val, System.Globalization.NumberStyles.Any, culture, out doubleVal) == false)
                {
                    return;
                }
            }
            Excel.Range r = DB.childs[nameL1].childs[nameL2.ToLower()].RangeByDay(day);
            r.Value = val;
            GlobalMethods.ToLog("Запись в Базу данных ячейка " + r.Address + " Субект {" + _name + "} " + nameL1 + " " + nameL2 + " день " + day + " значение " + val);
            //Marshal.ReleaseComObject(r);
        }
        public int GetCode(string name, string? L1 = null)
        {
            if (L1 == null) L1 = RangeReferences.ActiveL1;
            return DB.childs[L1].childs[name].Head.Column - DB.childs[L1].childs["основное"].Head.Column;
        }
        public void ChangeType(string oldName, string newName, string? L1 = null)
        {
            if (L1 == null)
            {
                L1 = PS._activeChild._name;
            }
            GlobalMethods.ToLog("Изменео {" + _name + "} " + L1 + " с '" + oldName + "' на '" + newName + "'");
            //Main.instance.stopped = true;
            Main.instance.StopAll();
            if (PS.childs[L1].childs.ContainsKey(newName))
            {
                PS.childs[L1].childs[oldName].Head.Value = newName;
                PS.childs[L1].childs[newName].Head.Value = oldName;
                PS.childs[L1].childs[oldName]._name = newName;
                PS.childs[L1].childs[newName]._name = oldName;

                ChildObject c1 = PS.childs[L1].childs[oldName];
                ChildObject c2 = PS.childs[L1].childs[newName];

                PS.childs[L1].childs.Remove(oldName);
                PS.childs[L1].childs.Remove(newName);

                PS.childs[L1].childs.Add(newName, c1);
                PS.childs[L1].childs.Add(oldName, c2);

                PS.childs[L1].childs[oldName].UpdateFormulas();
                PS.childs[L1].childs[oldName].UpdateColors();
            }
            else
            {
                PS.childs[L1].childs[oldName].Head.Value = newName;
                PS.childs[L1].childs[oldName]._name = newName;
                ChildObject c = PS.childs[L1].childs[oldName];
                PS.childs[L1].childs.Remove(oldName);
                PS.childs[L1].childs.Add(newName, c);

            }
            if (Main.instance.colors.mainTitle.ContainsKey(newName))
            {
                DB.childs[L1].childs["код"].Body.Cells[1, 1] = GetCode(newName, L1);
            }
            //Main.instance.stopped = false;
            Main.instance.ResumeAll();

            PS.childs[L1].childs[newName].UpdateFormulas();
            PS.childs[L1].childs[newName].UpdateColors();
        }
        public void ChangeTypeCell(string newName, string? L1 = null)
        {
            //if (L1 == null) L1 = 
        }
        public void Reset(string nameL1, string nameL2)
        {
            PS.childs[nameL1].childs[nameL2].UpdateFormulas();
            PS.childs[nameL1].childs[nameL2].UpdateColors();
            PS.childs[nameL1].childs[nameL2].ResetCode();
        }
        public void UpdateNames()
        {
            PS.UpdateNames();
        }
        public void UpdateAllColors()
        {
            if (PS != null)
            {
                PS.UpdateAllColors();
            }
        }
        public void UpdateBorders()
        {
            if (PS != null)
                PS.UpdateAllBorders();
            if (DB != null)
                DB.UpdateAllBorders();
        }
        public void UpdateAllPSFormulas()
        {
            PS.UpdateFormulas();
        }
        public void UpdateAllDBFormulas()
        {
            DB.UpdateFormulas();
        }
        public void UpdateReferencesPS()
        {
            PS.parentID = ID;
            PS.firstParentID = ID;
            PS.UpdateReferences();
            UpdateLevels();
            UpdateParents();
        }
        public void UpdateReferencesDB()
        {
            DB.parentID = ID;
            DB.firstParentID = ID;
            DB.UpdateReferences();
            UpdateLevels();
            UpdateParents();
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
            foreach (ChildObject item in childs.Values)
            {
                item.parentID = ID;
                item._name = _name;
                item.firstParentID = ID;
                item.UpdateParents();
            }
        }
        public void UpdateLevels()
        {
            foreach (ChildObject item in childs.Values)
            {
                item._level = Level.level0;
                item.UpdateLevels();
            }
        }
        public void AddNewPS(string nameL1, string nameL2, bool stopall = true)
        {
            PS.AddNewRange(nameL1, nameL2, stopall);
        }

        private void AddNewDBL1Standart(string nameL1, bool stopall = true)
        {
            DB.AddNewRange(nameL1, "код", stopall);
            DB.AddNewRange(nameL1, "основное", stopall);

            DB.childs[nameL1].childs["код"].UpdateFormulas(stopall);
            Excel.Range r = ((Excel.Range)DB.childs[nameL1].childs["код"].Body.Cells[1, 1]);
            r.Value = 1;
            Marshal.ReleaseComObject(r);

            DB.childs[nameL1].childs["основное"].UpdateFormulas(stopall);
        }
        public void AddNewDBL1StandartOther(string nameL1, bool stopall = true)
        {
            AddNewDBL1Standart(nameL1, stopall);
            DB.AddNewRange(nameL1, "корректировка факт", stopall);
            DB.AddNewRange(nameL1, "ручное", stopall);

            DB.childs[nameL1].childs["код"].UpdateFormulas(stopall);
            Excel.Range r = ((Excel.Range)DB.childs[nameL1].childs["код"].Body.Cells[1, 1]);
            r.Value = 1;
            Marshal.ReleaseComObject(r);

            DB.childs[nameL1].childs["основное"].UpdateFormulas(stopall);
        }

        public void CreateNewPS(string nameL1, string nameL2, bool stopall = true)
        {
            PS.CreateNewRange(nameL1, nameL2, stopall);
            PS.UpdateAllColors();
        }

        private void CreateNewDBL1Standart(string nameL1, bool stopall = true)
        {
            DB.CreateNewRange(nameL1, "код", stopall);
            DB.AddNewRange(nameL1, "основное", stopall);

            DB.childs[nameL1].childs["код"].UpdateFormulas(stopall);
            Excel.Range r = ((Excel.Range)DB.childs[nameL1].childs["код"].Body.Cells[1, 1]);
            r.Value = 1;
            Marshal.ReleaseComObject(r);

            DB.childs[nameL1].childs["основное"].UpdateFormulas(stopall);
        }
        public void CreateNewDBL1StandartOther(string nameL1, bool stopall = true)
        {
            CreateNewDBL1Standart(nameL1, stopall);
            DB.AddNewRange(nameL1, "корректировка факт", stopall);
            DB.AddNewRange(nameL1, "ручное", stopall);

            DB.childs[nameL1].childs["код"].UpdateFormulas(stopall);
            Excel.Range r = ((Excel.Range)DB.childs[nameL1].childs["код"].Body.Cells[1, 1]);
            r.Value = 1;
            Marshal.ReleaseComObject(r);

            DB.childs[nameL1].childs["основное"].UpdateFormulas(stopall);
        }

        public void AddDBPlansTable(bool stopall = true)
        {
            AddNewDBL1Standart("план", stopall);
            DB.AddNewRange("план", "заявка", stopall);
            DB.AddNewRange("план", "утвержденный", stopall);
            DB.AddNewRange("план", "корректировка", stopall);

            foreach (string n in DB.childs.Keys)
            {
                if (n != "план")
                {
                    DB.AddNewRange(n, "по плану", stopall);
                    DB.childs[n].childs["по плану"].UpdateFormulas(stopall);
                    // DB.childs[n].childs["по плану"].Body.Formula = "=" + ((Excel.Range)DB.childs["план"].childs["основное"].Body.Cells[1,1]).Address[false, false];
                }
            }
            DB.childs["план"].childs["заявка"].UpdateFormulas(stopall);
            DB.childs["план"].childs["код"].UpdateFormulas(stopall);
            // DB.childs["план"].childs["заявка"].Body.FormulaR1C1 = "=SUM(RC[1],RC[2])";
        }
        public void AddDBL2(string nameL1, string nameL2, bool stopall = true)
        {
            if (!DB.HasItem(nameL1)) AddNewDBL1StandartOther(nameL1, stopall);
            DB.AddNewRange(nameL1, nameL2, stopall);
        }

        public void AddPlans(bool stopall = true)
        {
            AddDBPlansTable(stopall);
            PS.AddNewRange("план", "заявка", stopall);
        }
        public void AddMeter(string nameL1, bool stopall = true)
        {
            DB.AddNewRange(nameL1, "счетчик", stopall);
            DB.AddNewRange(nameL1, "по счетчику", stopall);

            DB.childs[nameL1].childs["счетчик"].UpdateFormulas(stopall);
            DB.childs[nameL1].childs["по счетчику"].UpdateFormulas(stopall);

            UpdateMeterCoef();
        }
        public void UpdateMeterCoef()
        {
            if (meterCoef != null)
            {
                foreach (ChildObject item in DB.childs.Values)
                {
                    if (item.HasItem("по счетчику"))
                    {
                        Excel.Range r = ((Excel.Range)item.childs["по счетчику"].Body.Cells[1, 1]);
                        GlobalMethods.ToLog("Изменен коэфициент счетчика для {" + _name + "} с '" + r.Value + "' на '" + meterCoef + "'");
                        r.Value = meterCoef;
                        Marshal.ReleaseComObject(r);
                    }
                }
            }
        }


        public void RemovePlan()
        {
            GlobalMethods.ToLog("Удален план для {" + _name + "}");
            DB.childs["план"].Remove();
            if (PS.HasItem("план")) PS.childs["план"].Remove();

            foreach (ChildObject n in DB.childs.Values)
            {
                if (n.HasItem("по плану"))
                {
                    n.childs["по плану"].Remove();
                }
            }

            foreach (ChildObject n in PS.childs.Values)
            {
                
                if (n.HasItem("по плану"))
                {
                    ChangeType("по плану", "ручное", n._name);
                    n.ChangeCod();
                }
            }
        }
        public void RemoveMeter(string nameL1)
        {
            GlobalMethods.ToLog("Удалены по счетчику для {" + _name + "} " + nameL1);
            DB.RemoveRange(nameL1, "счетчик");
            DB.RemoveRange(nameL1, "по счетчику");
            if (!DB.HasItem("счетчик"))
            {
                meterCoef = null;
            }
        }

        public void ClearAll()
        {
            DB.ClearAll();
        }

        public void Check()
        {
            foreach (ChildObject item in childs.Values)
            {
                item.Check();
            }
        }

        public void ReleaseAllComObjects()
        {
            foreach (ChildObject item in childs.Values)
            {
                item.ReleaseAllComObjects();
            }
        }

        public void RemoveSubject()
        {
            GlobalMethods.ToLog("Субъект {" + _name + "} удален");

            string u0, u1, u2;
            u0 = ((Excel.Range)PS.Head.Offset[-3].MergeArea.Cells[1,1]).Value as string;
            u1 = ((Excel.Range)PS.Head.Offset[-2].MergeArea.Cells[1,1]).Value as string;
            u2 = ((Excel.Range)PS.Head.Offset[-1].MergeArea.Cells[1,1]).Value as string;

            if (Main.instance.heads.heads.ContainsKey(u0))
            {
                if (Main.instance.heads.heads[u0].childs.ContainsKey(u1))
                {
                    if (Main.instance.heads.heads[u0].childs[u1].childs.ContainsKey(u2))
                    {
                        if (Main.instance.heads.heads[u0].childs[u1].childs[u2].Range.Columns.Count == PS.Head.Columns.Count)
                        {
                            Main.instance.heads.heads[u0].childs[u1].childs.Remove(u2);
                        }
                    }
                    if (Main.instance.heads.heads[u0].childs[u1].Range.Columns.Count == PS.Head.Columns.Count)
                    {
                        Main.instance.heads.heads[u0].childs.Remove(u1);
                    }
                }
                if (Main.instance.heads.heads[u0].Range.Columns.Count == PS.Head.Columns.Count)
                {
                    Main.instance.heads.heads.Remove(u0);
                }
            }

            DB.Remove();
            PS.Remove();
            RangeReferences.idDictionary.Remove(ID);
            Main.instance.references.references.Remove(_name);
            ID = null;

            if (Main.instance.heads.heads.ContainsKey(u0))
            {
                if (Main.instance.heads.heads[u0].childs.ContainsKey(u1))
                {
                    if (Main.instance.heads.heads[u0].childs[u1].childs.ContainsKey(u2))
                    {
                        Main.instance.heads.heads[u0].childs[u1].childs[u2].UpdateBorders();
                    }
                    Main.instance.heads.heads[u0].childs[u1].UpdateBorders();
                }
                Main.instance.heads.heads[u0].UpdateBorders();
            }
        }
    
        public void UpdateHeads(bool stopall = true)
        {
            HeadObject h0 = HeadL0, h1 = HeadL1, h2 = HeadL2;
            if (h0.LastColumn.Column < PS.LastColumn.Column)
            {
                int resizeValue = PS.LastColumn.Column - h0.LastColumn.Column;
                GlobalMethods.ToLog("resizeValue H0: " + resizeValue);
                h0.Resize(resizeValue, false, stopall);
                h0.UpdateColors();
                h0.UpdateBorders();
            }
            
            if (h1.LastColumn.Column < PS.LastColumn.Column)
            {
                int resizeValue = PS.LastColumn.Column - h1.LastColumn.Column;
                GlobalMethods.ToLog("resizeValue H1: " + resizeValue);
                h1.Resize(resizeValue, false, stopall);
                h1.UpdateColors();
                h1.UpdateBorders();
            }

            if (h2.LastColumn.Column < PS.LastColumn.Column)
            {
                int resizeValue = PS.LastColumn.Column - h2.LastColumn.Column;
                GlobalMethods.ToLog("resizeValue H2: " + resizeValue);
                h2.Resize(resizeValue, false, stopall);
                h2.UpdateColors();
                h2.UpdateBorders();
            }
        }
    }
}
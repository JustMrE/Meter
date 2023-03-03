using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;
using Newtonsoft.Json;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace Meter
{
    public class ReferenceObject : ReferencesParent
    {
        public int? codPlan { get; set; }
        public string? meterCoef { get; set; }

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

        public int? ActiveDay()
        {
            return _activeChild._activeChild._activeChild.DayByRange(MenuBase._activeRange);
        }
        public bool HasRange(Excel.Range rng)
        {
            return Main.instance.xlApp.Intersect(PS.Range, rng) != null;
        }
        /*public void WriteToDB(int day, string val)
        {
            WriteToDB(_activeChild._activeChild._name, _activeChild._activeChild._activeChild._name.ToLower(), day, val);
        }*/
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
            Marshal.ReleaseComObject(r);
        }
        public void AddNewRange(string psdb, string nameL1, string nameL2 = "ручное")
        {

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
                PS.UpdateAllColors();
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
        public void AddNewPS(string nameL1, string nameL2)
        {
            PS.AddNewRange(nameL1, nameL2);
        }

        private void AddNewDBL1Standart(string nameL1)
        {
            DB.AddNewRange(nameL1, "код");
            DB.AddNewRange(nameL1, "основное");

            DB.childs[nameL1].childs["код"].UpdateFormulas();
            // DB.childs[nameL1].childs["код"].Body.Value = "=" + ((Excel.Range)DB.childs[nameL1].childs["код"].Body.Cells[1, 1]).Address;
            // DB.childs[nameL1].childs["код"].Body.Replace("$", "", LookAt: XlLookAt.xlPart);
            Excel.Range r = ((Excel.Range)DB.childs[nameL1].childs["код"].Body.Cells[1, 1]);
            r.Value = 1;
            Marshal.ReleaseComObject(r);

            DB.childs[nameL1].childs["основное"].UpdateFormulas();
            //DB.childs[nameL1].childs["основное"].Body.FormulaR1C1 = "=@INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) + RC[-1],4,1),TRUE)";
            //((Excel.Range)DB.childs[nameL1].childs["основное"].Body.Cells[1, 1]).Value = "";
        }
        public void AddNewDBL1StandartOther(string nameL1)
        {
            AddNewDBL1Standart(nameL1);
            // DB.AddNewRange(nameL1, "код");
            // DB.AddNewRange(nameL1, "основное");
            DB.AddNewRange(nameL1, "корректировка факт");
            DB.AddNewRange(nameL1, "ручное");

            // DB.childs[nameL1].childs["код"].Body.Value = "=" + ((Excel.Range)DB.childs[nameL1].childs["код"].Body.Cells[1, 1]).Address;
            // DB.childs[nameL1].childs["код"].Body.Replace("$", "", LookAt: XlLookAt.xlPart);
            DB.childs[nameL1].childs["код"].UpdateFormulas();
            Excel.Range r = ((Excel.Range)DB.childs[nameL1].childs["код"].Body.Cells[1, 1]);
            r.Value = 1;
            Marshal.ReleaseComObject(r);

            DB.childs[nameL1].childs["основное"].UpdateFormulas();
            // DB.childs[nameL1].childs["основное"].Body.FormulaR1C1 = "=@INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) + RC[-1],4,1),TRUE) + RC[1]";
            // ((Excel.Range)DB.childs[nameL1].childs["основное"].Body.Cells[1, 1]).Value = "";
        }
        public void AddDBPlansTable()
        {
            AddNewDBL1Standart("план");
            DB.AddNewRange("план", "заявка");
            DB.AddNewRange("план", "утвержденный");
            DB.AddNewRange("план", "корректировка");

            foreach (string n in DB.childs.Keys)
            {
                if (n != "план")
                {
                    DB.AddNewRange(n, "по плану");
                    DB.childs[n].childs["по плану"].UpdateFormulas();
                    // DB.childs[n].childs["по плану"].Body.Formula = "=" + ((Excel.Range)DB.childs["план"].childs["основное"].Body.Cells[1,1]).Address[false, false];
                }
            }
            DB.childs["план"].childs["заявка"].UpdateFormulas();
            DB.childs["план"].childs["код"].UpdateFormulas();
            // DB.childs["план"].childs["заявка"].Body.FormulaR1C1 = "=SUM(RC[1],RC[2])";
        }
        public void AddDBL2(string nameL1, string nameL2)
        {
            if (!DB.HasItem(nameL1)) AddNewDBL1StandartOther(nameL1);
            DB.AddNewRange(nameL1, nameL2);
        }

        public void AddPlans()
        {
            AddDBPlansTable();
            PS.AddNewRange("план", "заявка");
        }
        public void AddMeter(string nameL1)
        {
            DB.AddNewRange(nameL1, "счетчик");
            DB.AddNewRange(nameL1, "по счетчику");

            DB.childs[nameL1].childs["счетчик"].UpdateFormulas();
            DB.childs[nameL1].childs["по счетчику"].UpdateFormulas();

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
                        r.Value = meterCoef;
                        Marshal.ReleaseComObject(r);
                    }
                }
            }
        }


        public void RemovePlan()
        {
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

        public void Remove()
        {
            DB.Remove();
            PS.Remove();
            RangeReferences.idDictionary.Remove(ID);
            Main.instance.references.references.Remove(_name);
            ID = null;
        }
    }
}
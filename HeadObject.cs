using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;
using Newtonsoft.Json;
using System.Data.Common;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Security.Cryptography.Xml;

namespace Meter
{
    public class HeadObject
    {
        [JsonIgnore]
        private Excel.Range _range;
        [JsonIgnore]
        public Excel.Worksheet ws;

        public bool HasIndent(IndentDirection direction)
        { 
            Color? c1 ,c2;
            if (direction == IndentDirection.left)
            {
                try 
                {
                    c1 = ColorsData.GetRangeColor(FirstCell.Offset[0, -1]);
                    c2 = ColorsData.GetRangeColor(FirstCell.Offset[0, -2]);
                }
                catch
                {
                    c1 = null;
                    c2 = null;
                }
                if ((c1 != null && c2 != null) && (c1 == Color.White && c2  != Color.White))
                    return true;
                else
                    return false;
            }
            else
            {
                try 
                {
                    c1 = ColorsData.GetRangeColor(FirstCell.Offset[0, 1]);
                    c2 = ColorsData.GetRangeColor(FirstCell.Offset[0, 2]);
                }
                catch
                {
                    c1 = null;
                    c2 = null;
                }
                if ((c1 != null && c2 != null) && (c1 == Color.White && c2  != Color.White))
                    return true;
                else
                    return false;
            }
        }
        public string _name { get; set; }
        [JsonIgnore]
        public string? ID { get; set; }
        [JsonIgnore]
        public string? parentID { get; set; }
        public string WSName 
        {
            get
            {
                if (ws is null)
                    return null;
                else
                    return ws.CodeName;
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
                    return _range.Address;
            }
            set
            {   if (value is not null)
                    _range = WS.Range[value];
            }
        }

        public Level _level;

        [JsonIgnore]
        public Excel.Range Range 
        {
            get
            {
                try
                {
                    if (_range == null && rangeAddress != null)
                    {
                        _range = WS.Range[rangeAddress];
                    }
                    return _range;
                }
                catch (Exception)
                {

                    return null;
                }
                
            } 
            set
            {
                _range = value;
                rangeAddress = value.Address;
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
        public Excel.Range LastColumn
        {
            get
            {
                try
                {
                    return (Excel.Range)_range.Columns[_range.Columns.Count];
                }
                catch
                {
                    return null;
                }
            }
        }
        [JsonIgnore]
        public Excel.Range LastCell
        {
            get
            {
                try
                {
                    return (Excel.Range)_range.Cells[1, _range.Columns.Count];
                }
                catch
                {
                    return null;
                }
                
            }
        }
        [JsonIgnore]
        public Excel.Range FirstCell
        {
            get
            {
                try
                {
                    return (Excel.Range)_range.Cells[1, 1];
                }
                catch
                {
                    return null;
                }
                
            }
        }
        [JsonIgnore]
        public HeadObject GetParent
        {
            get
            {
                if (parentID != null && HeadReferences.idDictionary.ContainsKey(parentID))
                {
                    return HeadReferences.idDictionary[parentID];
                }
                else
                {
                    return null;
                }
            }
        }

        public HeadObject()
        {
            if (ID == null)
            {
                ID = Guid.NewGuid().ToString();
                while (HeadReferences.idDictionary.ContainsKey(ID))
                {
                    ID = Guid.NewGuid().ToString();
                }
            }
            
        }

        public Dictionary<string, HeadObject> childs { get; set; }

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

        public HeadObject HeadByRange(Excel.Range range)
        {
            HeadObject head;
            if (HasRange(range))
            {
                head = this;
            }
            else
            {
                if (childs != null)
                {
                    foreach (var child in childs.Values)
                    {
                        head = child.HeadByRange(range);
                        if (head != null)
                        {
                            return head;
                        }
                    }
                }
                head = null;
            }
            return head;
        }

        public HeadObject HeadByName(string name)
        {
            HeadObject head;
            if (name == _name)
            {
                head = this;
            }
            else
            {
                if (childs != null)
                {
                    foreach (var child in childs.Values)
                    {
                        head = child.HeadByName(name);
                        if (head != null)
                        {
                            return head;
                        }
                    }
                }
                head = null;
            }
            return head;
        }

        public void CreateChilds()
        {
            if (_level == Level.level2) return;
            if (_level == null) _level = Level.level0;
            int maxColumn = LastColumn.Column;
            int row = _range.Offset[1].Row;
            Excel.Range rng1 = ((Excel.Range)_range.Cells[1, 1]).Offset[1];
            while (rng1.Column <= maxColumn)
            {
                string name = (string)((Excel.Range)rng1.Cells[1, 1]).Value;
                if (string.IsNullOrEmpty(name))
                {
                    rng1 = rng1.Offset[0, 1];
                    continue;
                } 
                if (childs == null) childs = new Dictionary<string, HeadObject>();
                
                if (!childs.ContainsKey(name))
                {
                    childs.Add(name, new HeadObject()
                    {
                        _name = name,
                        WS = this.WS,
                        Range = rng1.MergeArea,
                        _level = this._level + 1,
                        parentID = ID,
                    });
                }
                rng1 = rng1.Offset[0, 1];
            }

            if (_level != Level.level2 && childs != null)
            {
                foreach (HeadObject c in childs.Values)
                {
                    c.CreateChilds();
                }
            }
        }

        public void Increase(int column, bool stopall = true)
        {
            if (stopall == true) Main.instance.StopAll();
            for (int i = 0; i < column; i++)
            {
                LastCell.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight);
            }
            if (stopall == true) Main.instance.ResumeAll();
        }

        public void Decrease(int column, bool stopall = true)
        {
            if (stopall == true) Main.instance.StopAll();
            for (int i = 0; i < column; i++)
            {
                LastCell.Delete(Shift: Excel.XlDeleteShiftDirection.xlShiftToLeft);
            }
            if (stopall == true) Main.instance.ResumeAll();
        }

        public void Resize(int column, bool newColumn = true, bool stopall = true)
        {
            if (newColumn)
            {
                Range = Range.Resize[Range.Rows.Count, column];
            }
            else
            {
                Range = Range.Resize[Range.Rows.Count, Range.Columns.Count + column];
            }

            if (stopall == true) Main.instance.xlApp.DisplayAlerts = false;
            Range.Merge();
            if (stopall == true) Main.instance.xlApp.DisplayAlerts = true;
        }
        public void UpdateColors()
        {
            if (_level == Level.level0)
            {
                Range.Interior.Color = Main.instance.colors.subColors["colorUp1"];
            }
            else if (_level == Level.level1)
            {
                Range.Interior.Color = Main.instance.colors.subColors["colorUp2"];
            }
            else if (_level == Level.level2)
            {
                Range.Interior.Color = Main.instance.colors.subColors["colorUp3"];
            }
        }
        public void UpdateBorders()
        {
            Range.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
        }
        public void UpdateAllColors()
        {
            UpdateColors();

            if (childs != null)
            {
                foreach (HeadObject c in childs.Values)
                {
                    c.UpdateAllColors();
                }
            }
        }
        public void UpdateParents()
        {
            if (!HeadReferences.idDictionary.ContainsKey(ID)) HeadReferences.idDictionary.Add(ID, this);
            if (childs != null)
            {
                foreach (HeadObject item in childs.Values)
                {
                    item.parentID = ID;
                    item.UpdateParents();
                }
            }
        }

        public void Increase(bool stopall = true)
        {
            Resize(1, false, stopall);
        }

        public void Decrese(bool stopall = true)
        {
            Resize(-1, false, stopall);
        }
    
        public void ReleaseAllComObjects()
        {
            if (childs != null)
            {
                foreach (HeadObject ho in childs.Values)
                {
                    ho.ReleaseAllComObjects();
                }
            }
            GlobalMethods.ReleseObject(Range);
            GlobalMethods.ReleseObject(WS);
            GlobalMethods.ReleseObject(LastCell);
            GlobalMethods.ReleseObject(LastColumn);
        }

        public void Indent(IndentDirection direction, bool stopall = true)
        {
            ChangeIndent(direction, stopall);
        }

        private void ChangeIndent(IndentDirection direction, bool stopall = true)
        {
            Excel.Range r;
                switch (direction)
                {
                    case IndentDirection.left :
                        r = FirstCell;
                        if (_level == Level.level0)
                        {
                            r = r.Offset[0, -1].Resize[42, 1];
                        }
                        else if (_level == Level.level1)
                        {
                            r = r.Offset[-1, -1].Resize[42, 1];
                        }
                        else if (_level == Level.level2)
                        {
                            r = r.Offset[-2, -1].Resize[42, 1];
                        }
                        break;
                    case IndentDirection.right :
                        r = LastCell;
                        if (_level == Level.level0)
                        {
                            r = r.Offset[0, 1].Resize[42, 1];
                        }
                        else if (_level == Level.level1)
                        {
                            r = r.Offset[-1, 1].Resize[42, 1];
                        }
                        else if (_level == Level.level2)
                        {
                            r = r.Offset[-2, 1].Resize[42, 1];
                        }
                        break;
                    default:
                        r = LastCell;
                        if (_level == Level.level0)
                        {
                            r = r.Offset[0, 1].Resize[42, 1];
                        }
                        else if (_level == Level.level1)
                        {
                            r = r.Offset[-1, 1].Resize[42, 1];
                        }
                        else if (_level == Level.level2)
                        {
                            r = r.Offset[-2, 1].Resize[42, 1];
                        }
                        break;
                }

                if (HasIndent(direction) == true)
                {
                    if (stopall) Main.instance.StopAll();
                    r.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                    if (stopall) Main.instance.ResumeAll();
                }
                else
                {
                    if (stopall) Main.instance.StopAll();
                    string adr = r.Address;
                    r.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                    r = Main.instance.wsCh.Range[adr];
                    if (_level == Level.level0)
                    {
                        //r = r.Offset[0, 1].Resize[42, 1];
                    }
                    else if (_level == Level.level1)
                    {
                        r = r.Offset[1].Resize[41, 1];
                    }
                    else if (_level == Level.level2)
                    {
                        r = r.Offset[2].Resize[40, 1];
                    }
                    r.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    r.Interior.Color = Color.White;

                    // ������ ������� ����� �������
                    r.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    r.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;

                    // ������ ������� ������ �������
                    r.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    r.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;

                    if (stopall) Main.instance.ResumeAll();
                }
                
        }

        public void Remove(bool stopall = true)
        {
            if (_level != Level.level0)
            {
                //if (GetParent.LastColumn.Column == LastColumn.Column)
                {
                    if (HasIndent(IndentDirection.right))
                    {
                        Indent(IndentDirection.right, stopall);
                    }
                    if (HasIndent(IndentDirection.left))
                    {
                        Indent(IndentDirection.left, stopall);
                    }
                }
            }
            if (_level != Level.level0)
                GetParent.childs.Remove(_name);
            else
                Main.instance.heads.heads.Remove(_name);
        }
    
        public void Delete(bool stopall = true)
        {
            if (stopall) Main.instance.StopAll();
            if (_level != Level.level2)
            {
                List<string> names = childs.Keys.ToList();
                foreach (string ho in names)
                {
                    childs[ho].Delete(false);
                }
                if (stopall) Main.instance.ResumeAll();
                return;
            }

            int rowOffset = 3 - (int)_level;
            Excel.Range r = FirstCell.Offset[rowOffset];
            string adr = r.Address[false, false];
            string name = (string)((Excel.Range)r.Cells[1,1]).Value;
            while (Range.Columns.Count > 1)
            {
                Main.instance.references.references[name].RemoveSubject(false);
                r = ((Excel.Worksheet)Main.instance.wb.ActiveSheet).Range[adr];
                name = (string)((Excel.Range)r.Cells[1,1]).Value;
            }
            Main.instance.references.references[name].RemoveSubject(false);
            if (stopall) Main.instance.ResumeAll();
        }
    
        public bool HasHead(string name)
        {
            bool hasName = false;
            if (_name == name)
                return true;
            if (childs != null)
            {
                foreach (HeadObject ho in childs.Values)
                {
                    hasName = ho.HasHead(name);
                    if (hasName == true)
                        return true;
                }
            }
            return false;
        }
    }
}
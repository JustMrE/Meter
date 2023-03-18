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

        [JsonIgnore]
        public bool indent { get; set; }
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
            Range.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
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
        public void UpdateIndents(HeadObject parent)
        {
            if (childs != null)
            {
                foreach (HeadObject item in childs.Values)
                {
                    item.UpdateIndents(this);
                }
            }

            if (parent != null)
            {
                if (parent.LastColumn.Column != LastColumn.Column)
                {
                    if (ColorsData.GetRangeColor(LastCell.Offset[0, 1]) == Color.White)
                    {
                        indent = true;
                    }
                    else
                    {
                        indent = false;
                    }
                }
                else
                { 
                    indent = false; 
                }
            }
            else
            {
                indent = false;
            }
        }

        public void Indent()
        {
            if (GetParent.LastColumn.Column == LastColumn.Column)
            {
                GetParent.Indent();
                indent = GetParent.indent;
            }
            else
            {
                Excel.Range r = LastCell;
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

                if (indent == true)
                {
                    indent = false;
                    Main.instance.StopAll();
                    r.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                    Main.instance.ResumeAll();
                }
                else
                {
                    indent = true;
                    Main.instance.StopAll();
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

                    // Задаем толстую левую границу
                    r.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    r.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;

                    // Задаем толстую правую границу
                    r.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    r.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;

                    Main.instance.ResumeAll();
                }
            }
        }

        public void Remove()
        {
            if (GetParent.LastColumn.Column == LastColumn.Column)
            {
                HeadObject ho = GetParent.HeadByRange(((Excel.Range)Range.Cells[1, 1]).Offset[0, -1]);
                if (ho != null)
                {
                    ho = GetParent.HeadByRange(((Excel.Range)Range.Cells[1, 1]).Offset[0, -2]);
                }
                if (ho != null)
                {
                    if (ho.indent == true)
                    {
                        ho.Indent();
                    }
                }
            }
            GetParent.childs.Remove(_name);
        }
    }
}
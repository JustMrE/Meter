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

        public string _name { get; set; }
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
                        _level = this._level + 1
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
    }
}
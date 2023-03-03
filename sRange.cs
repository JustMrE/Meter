using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;
using Newtonsoft.Json;

namespace Meter
{
    //public class sRange
    //{
    //    [JsonIgnore]
    //    public Excel.Range range;
    //    [JsonIgnore]
    //    public Excel.Worksheet ws;

    //    public string WSName 
    //    {
    //        get
    //        {
    //            if (ws is null)
    //                return null;
    //            else
    //                return ws.CodeName;
    //        }
    //        set
    //        {   if (value is not null)
    //            {
    //                if (value == "PS")
    //                {
    //                    ws = Main.instance.wsCh;
    //                }
    //                else if(value == "DB")
    //                {
    //                    ws = Main.instance.wsDb;
    //                }
    //            }
    //        }
    //    }
    //    public string Address 
    //    {
    //        get
    //        {
    //            if (range is null)
    //                return null;
    //            else
    //                return range.Address;
    //        }
    //        set
    //        {   if (value is not null)
    //                range = ws.Range[value];
    //        }
    //    }

    //    [JsonIgnore]
    //    public Excel.Range Range 
    //    {
    //        get
    //        {
    //            if (range == null && Address != null)
    //            {
    //                range = WS.Range[Address];
    //            }
    //            return range;
    //        } 
    //        set
    //        {
    //            range = value;
    //            Address = value.Address;
    //        }
    //    }
    //    [JsonIgnore]
    //    public Excel.Worksheet WS
    //    {
    //        get
    //        {
    //            if (ws == null && WSName != null)
    //            {
    //                if (WSName == "PS")
    //                {
    //                    ws = Main.instance.wsCh;
    //                }
    //                else if(WSName == "DB")
    //                {
    //                    ws = Main.instance.wsDb;
    //                }
    //            }
    //            return ws;
    //        } 
    //        set
    //        {
    //            ws = value;
    //            WSName = value.CodeName;
    //        }
    //    }
    //}
}
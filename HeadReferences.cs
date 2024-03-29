using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Newtonsoft.Json;

namespace Meter
{
    public class HeadReferences
    {
        public Dictionary<string, HeadObject> heads {get; set; }

        [JsonIgnore]
        public static Dictionary<string, HeadObject> idDictionary = new Dictionary<string, HeadObject>();

        public HeadReferences()
        {
            heads = new Dictionary<string, HeadObject>();
        }

        public HeadObject HeadByRange(Excel.Range range)
        {
            HeadObject head = null;
            foreach (var item in heads.Values)
            {
                head = item.HeadByRange(range);
                if (head != null)
                {
                    return head;
                }
            }
            return head;
        }

        public HeadObject HeadByName(string name)
        {
            HeadObject head = null;
            foreach (var item in heads.Values)
            {
                head = item.HeadByName(name);
                if (head != null)
                {
                    return head;
                }
            }
            return head;
        }

        public bool HasHead(string name)
        {
            bool hasHead = false;
            foreach (HeadObject ho in heads.Values)
            {
                hasHead = ho.HasHead(name);
                if (hasHead == true)
                    return true;
            }
            return false;
        }

        public void UpdateAllColors(bool stopall = true)
        {
            if (stopall) Main.instance.StopAll();
            foreach (HeadObject item in heads.Values)
            {
                item.UpdateAllColors();
            }
            if (stopall) Main.instance.ResumeAll();
        }

        public void UpdateParents()
        {
            foreach (HeadObject item in heads.Values)
            {
                item.UpdateParents();
            }
        }

        public void ReleaseAllComObjects()
        {
            heads.Values.AsParallel().ForAll(x => x.ReleaseAllComObjects());
        }
    }
}
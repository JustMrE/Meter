using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;
using Newtonsoft.Json;
using System.Runtime.InteropServices;

namespace Meter
{
    public class ReferencesParent
    {
        public string _name { get; set; }

        public string? ID;
        public Dictionary<string, ChildObject> childs;

        [JsonIgnore]
        public ChildObject _activeChild;

        public ReferencesParent() 
        {
            
        }
        public void ActivateTable(Excel.Range rng)
        {
            _activeChild = null;
            if (childs != null)
            {
                _activeChild = childs.Values.AsParallel().FirstOrDefault(n => n.HasRange(rng));
                if (_activeChild != null)
                {
                    _activeChild.ActivateTable(rng);
                    RangeReferences._activeObject = _activeChild;
                    //Marshal.ReleaseComObject(rng);
                }
            }
        }

        public virtual void UpdateBorders()
        {

        }
    
        public virtual void UpdateParents()
        {
            // if (ID == null || RangeReferences.idDictionary.ContainsKey(ID))
            // {
            //     ID = Guid.NewGuid().ToString();
            //     while (RangeReferences.idDictionary.ContainsKey(ID))
            //     {
            //         ID = Guid.NewGuid().ToString();
            //     }
            // }
        }

        
    }
}

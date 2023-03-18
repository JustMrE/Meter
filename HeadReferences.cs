using Main = Meter.MyApplicationContext;

namespace Meter
{
    public class HeadReferences
    {
        public Dictionary<string, HeadObject> heads {get; set; }

        public HeadReferences()
        {
            heads = new Dictionary<string, HeadObject>();
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

        public void ReleaseAllComObjects()
        {
            heads.Values.AsParallel().ForAll(x => x.ReleaseAllComObjects());
        }
    }
}
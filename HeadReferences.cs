namespace Meter
{
    public class HeadReferences
    {
        public Dictionary<string, HeadObject> heads {get; set; }

        public HeadReferences()
        {
            heads = new Dictionary<string, HeadObject>();
        }

        public void ReleaseAllComObjects()
        {
            foreach (HeadObject ho in heads.Values)
            {
                ho.ReleaseAllComObjects();
            }
        }
    }
}
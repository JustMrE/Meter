namespace Meter
{
    public class HeadReferences
    {
        public Dictionary<string, HeadObject> heads {get; set; }

        public HeadReferences()
        {
            heads = new Dictionary<string, HeadObject>();
        }
    }
}

using System.Collections.Concurrent;

public struct EMCOSObject
{
    public string name;
    public string? psid, dbid;
    // public int? emcosID;
    public ConcurrentDictionary<DateTime, string> values;
    public ConcurrentDictionary<DateTime, bool> flags;
}
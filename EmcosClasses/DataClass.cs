

namespace Emcos
{
    public class DataDatum
    {
        public int? ML_ID { get; set; }
        public int? POint_ID { get; set; }
        public int? ID { get; set; }
        public string? BT { get; set; }
        public string? ET { get; set; }
        public float? VAL { get; set; }
        public float? DR { get; set; }
        public object DF { get; set; }
        public string? READ_TIME { get; set; }
        public int? HSS { get; set; }
        public string? DSS { get; set; }
        public string? SFS { get; set; }
        public int? HAS_ACT { get; set; }
        public int? TFF_ID { get; set; }
        public int? BYP_EXISTS { get; set; }
    }

    public class DataRoot
    {
        public bool success { get; set; }
        public List<DataDatum> data { get; set; }
        public DateTime DB_TIME { get; set; }
        public int? bufferLen { get; set; }
        public string? buffer { get; set; }
    }
}


namespace Emcos
{
    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
    public class ListDatum
    {
        public int GR_ID { get; set; }
        public int ID { get; set; }
        public int? GRC_NR { get; set; }
        public object GRC_DESC { get; set; }
        public string GR_CODE { get; set; }
        public string GR_NAME { get; set; }
        public string USER_NAME { get; set; }
        public int USER_ID { get; set; }
        public string DBU { get; set; }
        public int IS_PUBLIC { get; set; }
        public string IS_PUBLIC_TXT { get; set; }
        public int GR_TYPE_ID { get; set; }
        public string GR_TYPE_NAME { get; set; }
        public string GR_TYPE_CODE { get; set; }
        public int PARENT { get; set; }
        public int HASCHILDS { get; set; }
        public object GRC_BT { get; set; }
        public string TYPE { get; set; }
        public object COLOR { get; set; }
        public object ADD_LABEL { get; set; }
        public int POINT_ID { get; set; }
        public object GRP_NR { get; set; }
        public object GRP_DESC { get; set; }
        public string POINT_NAME { get; set; }
        public string POINT_CODE { get; set; }
        public int POINT_ENABLED { get; set; }
        public string POINT_ENABLED_TXT { get; set; }
        public int POINT_COMMERCIAL { get; set; }
        public string POINT_COMMERCIAL_TXT { get; set; }
        public int POINT_INTERNAL { get; set; }
        public string POINT_INTERNAL_TXT { get; set; }
        public int POINT_AUTO_READ_ENABLED { get; set; }
        public string POINT_AUTO_READ_ENABLED_TXT { get; set; }
        public string ECP_NAME { get; set; }
        public object XA_ID { get; set; }
        public string POINT_TYPE_NAME { get; set; }
        public string POINT_TYPE_CODE { get; set; }
        public object MOU_BT { get; set; }
        public object MOU_ET { get; set; }
        public string METER_NUMBER { get; set; }
        public string METER_TYPE_NAME { get; set; }
        public object GRP_BT { get; set; }
    }

    public class ListRoot
    {
        public bool success { get; set; }
        public List<ListDatum> data { get; set; }
        public int bufferLen { get; set; }
        public string buffer { get; set; }
    }


}
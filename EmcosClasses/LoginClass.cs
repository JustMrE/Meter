

namespace Emcos
{
        public class Data
    {
        public User user { get; set; }
        public string token { get; set; }
        public List<string> privileges { get; set; }
    }

    public class LoginRoot
    {
        public bool success { get; set; }
        public Data data { get; set; }
        public int bufferLen { get; set; }
        public string buffer { get; set; }
    }

    public class User
    {
        public string uid { get; set; }
        public string username { get; set; }
        public object fullName { get; set; }
        public object company { get; set; }
        public object department { get; set; }
    }
}
using System.Net;
using System.Text;
using Meter;
using Newtonsoft.Json;

namespace Emcos
{
    public static class EmcosMethods 
    {
        private static string _uri = @"http://10.0.144.11:8080/ec3api/v1";
        private static string host = @"10.0.144.11:8080";
        private static HttpClient client = new HttpClient();
        private static string? username = null, password = null;
        private static string? token = null;

        static EmcosMethods()
        {
            #if DEBUG
            username = "a_bagdatova";
            password = "z123456";
            #endif
        }
        public static string Username 
        {
            get 
            {
                return username != null ? username : "";
            }
            set 
            {
                username = value;
            }
        }
        public static string Password 
        {
            get 
            {
                return password != null ? password : "";
            }
            set 
            {
                password = value;
            }
        }
        
        public static void SetUserPass(string username, string password)
        {
            Username = username;
            Password = password;
        }
        public static bool LoginToEmcos()
        {
            string uri = _uri + "/user/login";

            if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
            {
                string jsonRequest = string.Format("{{\"username\":\"{0}\",\"password\":\"{1}\"}}", username, password);
                var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");

                var responce = client.PostAsync(uri, content).Result;
                if (responce.IsSuccessStatusCode)
                {
                    var jsonResoponse = responce.Content.ReadAsStringAsync().Result;
                    LoginRoot loginResponce = JsonConvert.DeserializeObject<LoginRoot>(jsonResoponse);
                    token = loginResponce.data.token;
                    return true;
                }
            }
            return false;
        }
        public static float? GetValue(DateTime dataFrom, DateTime dataTo, ChildObject childObject)
        {
            float? value = null;
            int mlID = -1;
            if (childObject.emcosMLID == 385)
            {
                mlID = 1040;
            }
            else
            {
                mlID = 1041;
            }
            if (mlID != -1)
            {
                List<DataDatum> list = GetValues(dataFrom, dataTo, (int)childObject.emcosID, mlID);
                List<float?> vals = list.Select(v => v.VAL).ToList();
                if (vals.Count != 0)
                {
                    foreach (var val in vals)
                    {
                        if (value == null && val != null)
                        {
                            value = val;
                        }
                        else if ( value != null && val != null)
                        {
                            value += val;
                        }
                    }
                }
            }
            return value;
        }
        public static bool GetValue(DateTime dataFrom, DateTime dataTo, ChildObject childObject, ref float? value)
        {
            bool flag = false;
            int mlID = -1;
            if (childObject.emcosMLID == 385)
            {
                mlID = 1040;
            }
            else
            {
                mlID = 1041;
            }
            if (mlID != -1)
            {
                List<DataDatum> list = GetValues(dataFrom, dataTo, (int)childObject.emcosID, mlID);
                List<float?> vals = list.Select(v => v.VAL).ToList();
                if (vals.Count != 0)
                {
                    foreach (var val in vals)
                    {
                        if (!flag && val == null)
                        {
                            flag = true;
                        }
                        if (value == null && val != null)
                        {
                            value = val;
                        }
                        else if ( value != null && val != null)
                        {
                            value += val;
                        }
                    }
                }
            }
            return flag;
        }
        
        /*
        AGGS_ID=13 AGGS_NAME=15 минут
        AGGS_ID=4 AGGS_NAME=Час
        AGGS_ID=5 AGGS_NAME=Сутки

        MD_ID=12 MD_NAME=Пятнадцатиминутные
        MD_ID=4 MD_NAME=Часовые
        MD_ID=5 MD_NAME=Суточные

        ML_ID=1040 ML_NAME=A+ энергия за 15 минут   
        ML_ID=1041 ML_NAME=A- энергия за 15 минут
        ML_ID=3605 ML_NAME=A+ энергия за час
        ML_ID=3606 ML_NAME=A- энергия за час
        ML_ID=385 ML_NAME=A+ энергия за сутки
        ML_ID=386 ML_NAME=A- энергия за сутки
        */
        private static List<DataDatum> GetValues(DateTime dataFrom, DateTime dataTo, int pointID, int mlID)
        {
            string uri = _uri + "/archives/point";

            // string mlID = "[1040,1041]";
            int mdID = 12, aggsID = 13;
            string jsonRequest = string.Format("{{\"FROM\":\"{0}\",\"TO\":\"{1}\",\"POINT_ID\":{2},\"ML_ID\":[{3}],\"MD_ID\":{4},\"AGGS_ID\":{5},\"WO_BYP\":0,\"WO_ACTS\":0,\"BILLING_HOUR\":0,\"SHOW_MAP_DATA\":0,\"FREEZED\":1}}", dataFrom.ToString("yyyy-MM-dd"), dataTo.ToString("yyyy-MM-dd"), pointID, mlID, mdID, aggsID);
            
            HttpRequestMessage message = new HttpRequestMessage()
            {
                RequestUri = new Uri(uri),
                Method = HttpMethod.Post,
                Headers =
                {
                    {HttpRequestHeader.Host.ToString(), host},
                    {HttpRequestHeader.Connection.ToString(), "keep-alive"},
                    {HttpRequestHeader.ContentLength.ToString(), jsonRequest.Length.ToString()},
                    {HttpRequestHeader.Accept.ToString(),"appplication/json, text/plain, */*"},
                    {"st-token", token},
                    {HttpRequestHeader.ContentType.ToString(), "application/json;charset=UTF-8"},
                    {HttpRequestHeader.AcceptEncoding.ToString(), "gzip, deflate"}
                },
                Content = new StringContent(jsonRequest, Encoding.UTF8, "application/json")
            };
            var responce = client.SendAsync(message).Result;
            if (responce.IsSuccessStatusCode)
            {
                var jsonResoponse = responce.Content.ReadAsStringAsync().Result;
                DataRoot vals = JsonConvert.DeserializeObject<DataRoot>(jsonResoponse);
                return vals.data != null ? vals.data : null;
            }
            return null;
        }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Emcos;

namespace Meter
{
    public partial class EmcosPicker : Form
    {
        private static HttpClient client = new HttpClient();
        static Dictionary<string, Groups> emcosList = new Dictionary<string, Groups>();
        //static Dictionary<string, Groups> emcosList1 = new Dictionary<string, Groups>();
        static List<int> root = new List<int>();
        //static List<string> root1 = new List<string>();
        bool selectType;
        public int id {get; set; }
        // public int? MS_TYPE_ID { get; set; }
        // public int? AGGS_TYPE_ID { get; set; }
        // public int? MSF_ID { get; set; }
        // public int? AGGS_ID { get; set; }
        // public int? AGGF_ID { get; set; }
        // public int? DIR_ID { get; set; }
        // public int? MD_ID { get; set; }
        public int? ML_ID { get; set; }

        // public string nextType;
        
        private static string token = "";

        public EmcosPicker()
        {
            InitializeComponent();
        }

        private void EmcosPicker_Load(object sender, EventArgs e)
        {
            label1.Text = RangeReferences.activeTable._name + " " + RangeReferences.ActiveL1;
            root.Add(1);
            LoginToEmcos();
            GetGroup(1);
            listBox1.Items.AddRange(emcosList.Keys.ToArray());
        }

        protected void SelectNewGroup(object sender, EventArgs e)
        {
            if (selectType == false)
            {
                string name = (string)this.listBox1.SelectedItem;
                Groups g = emcosList[name];
                if (g.type == "GROUP")
                {
                    UpdateList(g.id);
                    root.Add(g.id);
                }
                else if (g.type == "POINT")
                {
                    id = g.id;
                    //selectType = true;
                    //GetGroup(id, "points");
                    //GetGroup(id, "signals");
                    //UpdateList1("MS_TYPE");
                    listBox1.Items.Clear();
                    listBox1.Items.Add("A+ энергия за сутки");
                    listBox1.Items.Add("A- энергия за сутки");
                    selectType = true;
                }
            }
            else
            {
                if ((string)this.listBox1.SelectedItem == "A- энергия за сутки")
                {
                    ML_ID = 386;
                }
                else if ((string)this.listBox1.SelectedItem == "A+ энергия за сутки")
                {
                    ML_ID = 385;
                }
                DialogResult = DialogResult.OK;
                Close();
            }
            /*else
            {
                string name = (string)this.listBox1.SelectedItem;
                Groups g = emcosList1[name];
                if (g.type != "ml_from_t0")
                {
                    UpdateList1(nextType);
                    root1.Add(nextType);
                }
                else if (g.type == "ml_from_t0")
                {
                    UpdateList1(nextType);
                    DialogResult = DialogResult.OK;
                    Close();
                }
            }*/
            
        }

        private void UpdateList(int id)
        {
            GetGroup(id);
            if (emcosList.Keys.Count == 0) GetGroup(id, "points");
            listBox1.Items.Clear();
            listBox1.Items.AddRange(emcosList.Keys.ToArray());
        }

        /*private void UpdateList1(string TYPE)
        {
            string jsonRequest ="";
            string type = TYPE.ToLower();

            if (TYPE == "MS_TYPE")
            {
                jsonRequest = "{\"GR_ID\":[],\"POINT_ID\":[" + id + "],\"EXTRACTBRANCHDATA\":1,\"SCOPE\":\"ARCHIVES\"}";
                nextType = "AGGS_TYPE";
            }
            else if (TYPE == "AGGS_TYPE")
            {
                jsonRequest = "{\"GR_ID\":[],\"POINT_ID\":[" + id + "],\"EXTRACTBRANCHDATA\":1,\"SCOPE\":\"ARCHIVES\",\"MS_TYPE_ID\":" + MS_TYPE_ID + "}";
                nextType = "MSF";
            }
            else if (TYPE == "MSF")
            {
                jsonRequest = "{\"GR_ID\":[],\"POINT_ID\":[" + id + "],\"EXTRACTBRANCHDATA\":1,\"SCOPE\":\"ARCHIVES\",\"AGGS_TYPE_ID\":" + AGGS_TYPE_ID + ",\"MS_TYPE_ID\":" + MS_TYPE_ID + "}";
                nextType = "ML_T0";
            }
            else if (TYPE == "ML_T0")
            {
                type = "ml";
                jsonRequest = "{\"GR_ID\":[],\"POINT_ID\":[" + id + "],\"EXTRACTBRANCHDATA\":1,\"SCOPE\":\"ARCHIVES\",\"AGGS_TYPE_ID\":" + AGGS_TYPE_ID + ",\"MSF_ID\":" + MSF_ID + "}";
                nextType = "ML";
            }
            else if (TYPE == "ML")
            {
                type = "ml_from_t0";
                jsonRequest = "{\"GR_ID\":[],\"POINT_ID\":[" + id + "],\"EXTRACTBRANCHDATA\":1,\"SCOPE\":\"ARCHIVES\",\"MSF_ID\":" + MSF_ID + ",\"AGGS_ID\":" + AGGS_ID + ",\"AGGF_ID\":" + AGGF_ID + ",\"DIR_ID\":" + DIR_ID + ",\"MD_ID\":" + MD_ID + "}";
            }

            Test(jsonRequest, type);
            //if (emcosList1.Keys.Count == 0) GetType(id, "points");
            listBox1.Items.Clear();
            listBox1.Items.AddRange(emcosList1.Keys.ToArray());
        }*/

        private void btnOk_Click(object sender, EventArgs e)
        {

        }
        private void btnBack_Click(object sender, EventArgs e)
        {
            //if (selectType == false)
            {
                int last = root.Count - 1;
                int id = root[last - 1];
                if (last > 0)
                    root.RemoveAt(last);
                UpdateList(id);
            }
            /*else
            {
                int last = root1.Count - 1;
                if (last > 0)
                {
                    root1.RemoveAt(last);
                    string id = root1[last - 1];
                    UpdateList1(id);
                }
            }*/
            
        }
        private static void LoginToEmcos()
        {
            string uri = @"http://10.0.144.11:8080/ec3api/v1/user/login";
            string host = @"10.0.144.11:8080";
            string jsonRequest = "{\"username\":\"a_bagdatova\",\"password\":\"z123456\"}";
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
                    {"st-token", "undefined"},
                    {HttpRequestHeader.ContentType.ToString(), "application/json;charset=UTF-8"},
                    {HttpRequestHeader.AcceptEncoding.ToString(), "gzip, deflate"}
                },
                Content = new StringContent(jsonRequest, Encoding.UTF8, "application/json")
            };
            using (var responce = client.SendAsync(message).Result)
            {
                var jsonResoponse = responce.Content.ReadAsStringAsync().Result;
                var data = (JObject)JsonConvert.DeserializeObject(jsonResoponse);

                LoginRoot loginResponce = JsonConvert.DeserializeObject<LoginRoot>(jsonResoponse);
                token = loginResponce.data.token;
            }
        }

        private void GetGroup(int id, string type = "groups")
        {
            string uri = "http://10.0.144.11:8080/ec3api/v1/pointmenuv2/" + type + "?gr_id=" + id;
            string host = @"10.0.144.11:8080";
            HttpRequestMessage message = new HttpRequestMessage()
            {
                RequestUri = new Uri(uri),
                Method = HttpMethod.Get,
                Headers = 
                {
                    {HttpRequestHeader.Host.ToString(), host},
                    {HttpRequestHeader.Connection.ToString(), "keep-alive"},
                    {HttpRequestHeader.Accept.ToString(),"appplication/json, text/plain, */*"},
                    {"st-token", token},
                    {HttpRequestHeader.ContentType.ToString(), "application/json;charset=UTF-8"},
                    {HttpRequestHeader.AcceptEncoding.ToString(), "gzip, deflate"}
                },
            };
            using (var responce = client.SendAsync(message).Result)
            {
                var jsonResoponse = responce.Content.ReadAsStringAsync().Result;
                ListRoot vals = JsonConvert.DeserializeObject<ListRoot>(jsonResoponse);
                emcosList.Clear();
                foreach (var v in vals.data)
                {
                    string name = "";
                    if (v.TYPE == "POINT")
                    {
                        name = v.POINT_NAME;
                    }
                    else if (v.TYPE == "GROUP")
                    {
                        name = v.GR_NAME;
                    }
                    emcosList.Add(name, new Groups(){
                        name = name,
                        id = v.ID,
                        type = v.TYPE
                    });
                }
            }
        }
        
        // private void Test(string jsonRequest, string type)
        // {
        //     var client = new HttpClient();
        //     var request = new HttpRequestMessage(HttpMethod.Post, "http://10.0.144.11:8080/ec3api/v1/parammenuv2/" + type);
        //     request.Headers.Add("Connection", "keep-alive");
        //     request.Headers.Add("Accept", "application/json, text/plain, */*");
        //     request.Headers.Add("Accept-Language", "ru_RU");
        //     request.Headers.Add("st-token", "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJodHRwOlwvXC9zaWdtYXRlbGFzLmx0IiwiYXVkIjoiaHR0cDpcL1wvc2lnbWF0ZWxhcy5sdCIsImlhdCI6MTY3Njg3NDc1MiwibmJmIjoxNjc2ODc0NzUyLCJleHAiOjE2NzY4ODE5NTIsImRhdGEiOnsidWlkIjoiMTI3IiwidXNlcm5hbWUiOiJhX2JhZ2RhdG92YSIsIm5hbWUiOiJBX0JBR0RBVE9WQSIsImlzRXh0ZXJuYWwiOmZhbHNlLCJrZXkiOiJcLzJUY2xpSm1rRlJwaUdoaTNQXC9lMEE9PSJ9fQ.20AJRIwORJTDOPI2Z8fY_iyNs0WKNlNqIBCegQSk-VY");
        //     //request.Headers.Add("Origin", "http://10.0.144.11:8080");
        //     //request.Headers.Add("Referer", "http://10.0.144.11:8080/emcos3/archives");
        //     request.Headers.Add("Accept-Encoding", "gzip, deflate");
        //     var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");
        //     request.Content = content;
        //     //var response = await client.SendAsync(request);
        //     //response.EnsureSuccessStatusCode();
        //     //string s = await response.Content.ReadAsStringAsync();
        //     using (var responce = client.SendAsync(request).Result)
        //     {
        //         var jsonResoponse = responce.Content.ReadAsStringAsync().Result;
        //         ListRoot vals = JsonConvert.DeserializeObject<ListRoot>(jsonResoponse);
        //         emcosList1.Clear();
        //         foreach (var v in vals.data)
        //         {
        //             string name = "";
        //             if (v.TYPE == "MS_TYPE")
        //             {
        //                 name = v.MS_TYPE_NAME;
        //                 MS_TYPE_ID = v.MS_TYPE_ID;
        //             }
        //             else if (v.TYPE == "AGGS_TYPE")
        //             {
        //                 name = v.AGGS_TYPE_NAME;
        //                 AGGS_TYPE_ID = v.AGGS_TYPE_ID;
        //             }
        //             else if (v.TYPE == "MSF")
        //             {
        //                 name = v.MSF_NAME;
        //                 MSF_ID = v.MSF_ID;
        //             }
        //             else if (v.TYPE == "ML_T0")
        //             {
        //                 name = v.ML0_NAME;
        //                 AGGS_ID = v.AGGS_ID;
        //                 AGGF_ID = v.AGGF_ID;
        //                 DIR_ID = v.DIR_ID;
        //                 MD_ID = v.MD_ID;
        //             }
        //             else if (v.TYPE == "ML")
        //             {
        //                 name = v.ML_NAME;
        //                 ML_ID = v.ML_ID;
        //             }
        //             emcosList1.Add(name, new Groups(){
        //                 name = name,
        //                 type = v.TYPE,
        //             });
        //         }
        //     }
        //     /*var jsonResoponse = response.Content.ReadAsStringAsync().Result;
        //     ListRoot vals = JsonConvert.DeserializeObject<ListRoot>(jsonResoponse);
        //     emcosList1.Clear();
        //     foreach (var v in vals.data)
        //     {
        //         string name = "";
        //         if (v.TYPE == "MS_TYPE")
        //         {
        //             name = v.MS_TYPE_NAME;
        //             MS_TYPE_ID = v.MS_TYPE_ID;
        //         }
        //         else if (v.TYPE == "AGGS_TYPE")
        //         {
        //             name = v.AGGS_TYPE_NAME;
        //             AGGS_TYPE_ID = v.AGGS_TYPE_ID;
        //         }
        //         else if (v.TYPE == "MSF")
        //         {
        //             name = v.MSF_NAME;
        //             MSF_ID = v.MSF_ID;
        //         }
        //         else if (v.TYPE == "ML_T0")
        //         {
        //             name = v.ML0_NAME;
        //             AGGS_ID = v.AGGS_ID;
        //             AGGF_ID = v.AGGF_ID;
        //             DIR_ID = v.DIR_ID;
        //             MD_ID = v.MD_ID;
        //         }
        //         else if (v.TYPE == "ML")
        //         {
        //             name = v.ML_NAME;
        //             ML_ID = v.ML_ID;
        //         }
        //         emcosList1.Add(name, new Groups(){
        //             name = name,
        //             type = v.TYPE,
        //         });
        //     }*/
        // }
    }
}

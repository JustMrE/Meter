using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using Emcos;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;
using System.Collections.Concurrent;
using System.Globalization;

namespace Meter.Forms
{
    public partial class NewMenu : Meter.Forms.NewMenuBase
    {
        bool changed;

        public NewMenu()
        {
            InitializeComponent();
        }

        protected override void NewMenuBase_Shown(object sender, EventArgs e)
        {
            base.NewMenuBase_Shown(sender, e);
            textBox1.Text = DateTime.Today.AddDays(-1).ToString("dd");
        }
        public override void FormClose()
        {
            base.FormClose();
        }
         protected override void TextBox1_TextChanged(object sender, EventArgs e)
        {
            base.TextBox1_TextChanged(sender, e);
            string tbVal = textBox1.Text;
            if (string.IsNullOrEmpty(tbVal) || changed)
            {
                return;
            }

            int val = Int32.Parse(tbVal);
            if (val > 31)
            {
                changed = true;
                textBox1.Text = "31";
                changed = false;
            }
        
        }
        protected override void TextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if (!Char.IsDigit(c) && c != 8)
            {
                e.Handled = true;
            }
            if (textBox1.Text.Length > 2 && c != 8)
            {
                e.Handled = true;
            }
        }

        #region  Buttons
        protected override void Button2_Click(object sender, EventArgs e)
        {
            base.Button2_Click(sender, e);
            CultureInfo provider = CultureInfo.CreateSpecificCulture("ru-RU");
            ReferenceObject[] ranges = Main.instance.references.references.Values.Where(n => n.HasEmcosID == true).ToArray();
            Login();
            string format = "dd MMMM yyyy";
            string data = "2023-02-" + this.textBox1.Text;
            data = this.textBox1.Text.PadLeft(2,'0') + " " + this.lblMonth.Text + " " + this.lblYear.Text;
            DateTime result;
            DateTime.TryParseExact(data, format, provider, DateTimeStyles.None, out result);
            data = result.ToString("yyyy-MM-dd");
            if (string.IsNullOrEmpty(this.textBox1.Text) ||  (Int32.Parse(this.textBox1.Text) <= 0 && Int32.Parse(this.textBox1.Text) > 31))
            {
                MessageBox.Show("Не введена дата записи!");
                return;
            }
            
            ConcurrentDictionary<string, string> emcosValues = new ConcurrentDictionary<string, string>();
            //foreach (ReferenceObject item in ranges)
            Parallel.ForEach(ranges, item => 
            {
                foreach (var v in item.DB.childs.Values)
                {
                    if (v.HasItem("аскуэ") && v.emcosID != null)
                    {
                        //string q = v.Read("аскуэ", 0);
                        int id = (int)v.emcosID;
                        float? floatVal = GetValue(data, data, v.emcosID.ToString()).data.Where(n => n.ML_ID == v.emcosMLID).FirstOrDefault().VAL;
                        string? val = null;
                        if (floatVal != null)
                        {
                            floatVal = floatVal / 1000f;
                            val = floatVal.ToString().Replace(",", ".");
                        }
                        else
                        {
                            val = "0";
                        }
                        emcosValues.TryAdd(item.DB.childs[v._name].childs["аскуэ"].ID, val);
                    }
                }
            });

            Main.instance.StopAll();
            foreach (string id in emcosValues.Keys)
            {
                // ((ChildObject)RangeReferences.idDictionary[id]).RangeByDay(int.Parse(this.textBox1.Text)).Value = emcosValues[id];
                ((ChildObject)RangeReferences.idDictionary[id]).WriteValue(int.Parse(this.textBox1.Text), emcosValues[id]);
            }
            Main.instance.ResumeAll();
            MessageBox.Show("Done!");
        }
        protected override void btnAdmin_Click(object sender, EventArgs e)
        {
            base.btnAdmin_Click(sender, e);
            using (Password pass = new Password())
            {
                if (pass.ShowDialog() == DialogResult.OK)
                {
                    //if (pass.Val == "123456")
                    {
                        GlobalMethods.ToLog("Введен правильный пароль");
                        Main.instance.menu = Main.menues[1] as NewMenuBase;
                        Main.instance.menu.Show();
                        this.Hide();
                        GlobalMethods.CalculateFormsPositions();
                    }
                }
            }
        }
        #endregion

        public override void ActivateSheet(object sh)
        {
            // if (((Excel.Worksheet)Main.instance.xlApp.ActiveSheet).CodeName == "PS")
            if (((Excel.Worksheet)sh).CodeName == "PS")
                base.ActivateSheet(sh);
        }
        public override void DeactivateSheet()
        {
            base.DeactivateSheet();
        }
        protected override void NewMenuBase_Activated(object sender, EventArgs e)
        {
            Main.instance.menu = this;
            base.NewMenuBase_Activated(sender, e);
        }
        // public static void Login()
        // {
        //     using (HttpClient client = new HttpClient())
        //     {
        //         string uri = @"http://10.0.144.11:8080/ec3api/v1/user/login";
        //         string host = @"10.0.144.11:8080";
        //         string jsonRequest = "{\"username\":\"a_bagdatova\",\"password\":\"z123456\"}";
        //         HttpRequestMessage message = new HttpRequestMessage()
        //         {
        //             RequestUri = new Uri(uri),
        //             Method = HttpMethod.Post,
        //             Headers = 
        //             {
        //                 {HttpRequestHeader.Host.ToString(), host},
        //                 {HttpRequestHeader.Connection.ToString(), "keep-alive"},
        //                 {HttpRequestHeader.ContentLength.ToString(), jsonRequest.Length.ToString()},
        //                 {HttpRequestHeader.Accept.ToString(),"appplication/json, text/plain, */*"},
        //                 {"st-token", "undefined"},
        //                 {HttpRequestHeader.ContentType.ToString(), "application/json;charset=UTF-8"},
        //                 {HttpRequestHeader.AcceptEncoding.ToString(), "gzip, deflate"}
        //             },
        //             Content = new StringContent(jsonRequest, Encoding.UTF8, "application/json")
        //         };
        //         using (var responce = client.SendAsync(message).Result)
        //         {
        //             var jsonResoponse = responce.Content.ReadAsStringAsync().Result;
        //             var data = (JObject)JsonConvert.DeserializeObject(jsonResoponse);

        //             LoginRoot loginResponce = JsonConvert.DeserializeObject<LoginRoot>(jsonResoponse);
        //             token = loginResponce.data.token;
        //         }
        //     }
        // }
        // public static DataRoot GetValue(string dataFrom, string dataTo, string ID)
        // {
        //     using (HttpClient client = new HttpClient())
        //     {
        //         string uri = @"http://10.0.144.11:8080/ec3api/v1/archives/point";
        //         string host = @"10.0.144.11:8080";
        //         string jsonRequest = "{\"FROM\":\"" + dataFrom +"\",\"TO\":\""+ dataTo + "\",\"POINT_ID\":" + ID + ",\"ML_ID\":[385,386],\"MD_ID\":5,\"AGGS_ID\":5,\"WO_BYP\":0,\"WO_ACTS\":0,\"BILLING_HOUR\":0,\"SHOW_MAP_DATA\":0,\"FREEZED\":1}";
        //         HttpRequestMessage message = new HttpRequestMessage()
        //         {
        //             RequestUri = new Uri(uri),
        //             Method = HttpMethod.Post,
        //             Headers = 
        //             {
        //                 {HttpRequestHeader.Host.ToString(), host},
        //                 {HttpRequestHeader.Connection.ToString(), "keep-alive"},
        //                 {HttpRequestHeader.ContentLength.ToString(), jsonRequest.Length.ToString()},
        //                 {HttpRequestHeader.Accept.ToString(),"appplication/json, text/plain, */*"},
        //                 {"st-token", token},
        //                 {HttpRequestHeader.ContentType.ToString(), "application/json;charset=UTF-8"},
        //                 {HttpRequestHeader.AcceptEncoding.ToString(), "gzip, deflate"}
        //             },
        //             Content = new StringContent(jsonRequest, Encoding.UTF8, "application/json")
        //         };
        //         DataRoot vals;
        //         using (var responce = client.SendAsync(message).Result)
        //         {
        //             var jsonResoponse = responce.Content.ReadAsStringAsync().Result;
        //             vals = JsonConvert.DeserializeObject<DataRoot>(jsonResoponse);
        //         }
        //         return vals;
        //     }
        // }
    }
}

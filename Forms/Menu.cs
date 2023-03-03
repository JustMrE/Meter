using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Main = Meter.MyApplicationContext;
using System.Net;
using Emcos;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Text;
using System.Runtime.InteropServices;


namespace Meter
{
    public partial class Menu : MenuBase
    {
        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOOLWINDOW = 0x80;

        bool changed;

        public Menu()
        {
            InitializeComponent();
            this.Load += new System.EventHandler(this.Menu_Load);
            this.Shown += new EventHandler(this.Menu_Show);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(Menu_Close);
            Main.instance.menu = this;

            SetWindowLong(this.Handle, GWL_EXSTYLE, WS_EX_TOOLWINDOW);
        }
        protected override void Menu_Show(object sender, EventArgs e)
        {
            base.Menu_Show(sender, e);
            textBox1.Text = DateTime.Today.AddDays(-1).ToString("dd");
        }
        
        public override void ContextMenu()
        {
            base.ContextMenu();
        }
        protected override void RepairMenu_Click(object sender, EventArgs e)
        {
            base.RepairMenu_Click(sender, e);
        }
        protected void Button2_Click(object sender, EventArgs e)
        {
            ReferenceObject[] ranges = Main.instance.references.references.Values.Where(n => n.HasEmcosID == true).ToArray();
            Login();
            string data = "2023-02-" + this.textBox1.Text;
            if (string.IsNullOrEmpty(this.textBox1.Text) ||  (Int32.Parse(this.textBox1.Text) <= 0 && Int32.Parse(this.textBox1.Text) > 31))
            {
                MessageBox.Show("Не введена дата записи!");
                return;
            }
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
                        item.WriteToDB( v._name, "аскуэ", int.Parse(this.textBox1.Text), val);
                    }
                }
            });
            MessageBox.Show("Done!");
        }
        protected override void btnAdmin_Click(object sender, EventArgs e)
        {
            using (Password pass = new Password())
            {
                if (pass.ShowDialog() == DialogResult.OK)
                {
                    //if (pass.Val == "123456")
                    {
                        Main.forms[1].Show();
                        Main.instance.menu = (MenuBase)Main.forms[1];
                        this.Hide();
                        GlobalMethods.CalculateFormsPositions();
                    }
                }
            }
        }
        protected override void TextBox1_TextChanged(object sender, EventArgs e)
        {
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
        private void TextBox1_KeyPress(object sender, KeyPressEventArgs e)
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
        public override void ActivateSheet(object sh)
        {
            if (((Excel.Worksheet)Main.instance.xlApp.ActiveSheet).CodeName == "PS")
            base.ActivateSheet(sh);
        }
        public override void DeactivateSheet()
        {
            base.DeactivateSheet();
        }

        static void Login()
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
        static DataRoot GetValue(string dataFrom, string dataTo, string ID)
        {
            string uri = @"http://10.0.144.11:8080/ec3api/v1/archives/point";
            string host = @"10.0.144.11:8080";
            string jsonRequest = "{\"FROM\":\"" + dataFrom +"\",\"TO\":\""+ dataTo + "\",\"POINT_ID\":" + ID + ",\"ML_ID\":[385,386],\"MD_ID\":5,\"AGGS_ID\":5,\"WO_BYP\":0,\"WO_ACTS\":0,\"BILLING_HOUR\":0,\"SHOW_MAP_DATA\":0,\"FREEZED\":1}";
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
            DataRoot vals;
            using (var responce = client.SendAsync(message).Result)
            {
                var jsonResoponse = responce.Content.ReadAsStringAsync().Result;
                vals = JsonConvert.DeserializeObject<DataRoot>(jsonResoponse);
            }
            return vals;
        }
    }
}
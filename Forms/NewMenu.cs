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
        protected override void Button3_Click(object sender, EventArgs e)
        {
            DateTime date;
            DateTime.TryParseExact(textBox1.Text + " " + lblMonth.Text + " " + lblYear.Text, "dd MMMM yyyy", GlobalMethods.culture, DateTimeStyles.None, out date);
            string datstr = date.ToString("dd.MM.yy");
            base.Button3_Click(sender, e);

            if (Main.instance.wsTEPm.Range["A6"].Value == "")
            {
                WriteTEPs(date);
            }
            else
            {
                string lastDateStr = (string)Main.instance.wsTEPm.Range["A6"].Value;
                DateTime lastDate;
                DateTime.TryParseExact(lastDateStr, "dd.MM.yy", GlobalMethods.culture, DateTimeStyles.None, out lastDate);
                if (lastDate != date && lastDate < date && (date - lastDate).TotalDays == 1)
                {
                    WriteMaketTEP(date);
                    WriteTEP(date);
                }
                else if (lastDate == date || lastDate > date)
                {
                    if (MessageBox.Show("Запись на дату " + date.ToString("dd.MM.yy") + " уже существует!\nВы хотите произвести повторную запись?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        WriteTEPs(date, true);
                    }
                    else
                    {
                        return;
                    }
                }
                else if (lastDate < date && (date - lastDate).TotalDays > 1)
                {
                    MessageBox.Show("Нет записи за прошлые дни!\nВы Пытаетесь записать на " + date.ToString("dd.MM.yy") + " а последняя запись производилась на " + lastDate.ToString("dd.MM.yy"));
                    return;
                }

            }
            MessageBox.Show("Done!");
        }
        protected override void Button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Запись формы W89");
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
        
        private void WriteTEPs(DateTime date, bool rewrite = false)
        {
            WriteMaketTEP(date);
            WriteTEP(date, rewrite);
            CreateTXTforCDU();
        }
        private void WriteTEP(DateTime date, bool rewrite = false)
        {
            Main.instance.StopAll();
            string datstr = date.ToString("dd.MM.yy");
            int? row;
            GlobalMethods.ToLog("Зписываются ТЭПм и ТЭПн на " + datstr);

            if (rewrite == false)
            {
                Main.instance.wsTEPm.Range["6:6"].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                Main.instance.wsTEPm.Range["A6"].Value = datstr;
                Main.instance.wsTEPn.Range["6:6"].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                Main.instance.wsTEPn.Range["A6"].Value = datstr;
                row = 6;
            }
            else
            {
                try
                {
                    row = Main.instance.wsTEPn.Range["A:A"].Find(What: datstr, LookAt: Excel.XlLookAt.xlWhole).Row;
                }
                catch
                {
                    row = null;
                }
            }
            if (row != null)
            {
                int day = int.Parse(this.textBox1.Text);
                List<ChildObject> coList = Main.instance.references.references.Values.SelectMany(n => n.PS.childs.Values).Where(m => m.codTEP != null).ToList();
                foreach (ChildObject co in coList)
                {
                    co.WriteToTEP(day, (int)row, rewrite: rewrite);
                }
            }
            Main.instance.ResumeAll();
        }
        private void WriteMaketTEP(DateTime date)
        {
            string datstr = date.ToString("dd.MM.yy");
            GlobalMethods.ToLog("Зписываются ТЭПм и ТЭПн на " + datstr);

            Main.instance.StopAll();
            Main.instance.wsMTEP.Range["B:B"].ClearContents();
            Main.instance.ResumeAll();
            Main.instance.wsMTEP.Range["B1"].Value = datstr;
            int day = int.Parse(this.textBox1.Text);
            List<ChildObject> coList = Main.instance.references.references.Values.SelectMany(n => n.PS.childs.Values).Where(m => m.codMaketTEP != null).ToList();
            foreach (ChildObject co in coList)
            {
                co.WriteToMaketTEP(day);
            }
        }
        private void CreateTXTforCDU()
        {
            string txtForCDU = "";
            txtForCDU += Main.instance.wsMTEP.Range["B1"].Value + "\n";

            Excel.Range r1 = Main.instance.wsMTEP.Range["A2"];
            Excel.Range r2 = r1.End[Excel.XlDirection.xlDown];
            Excel.Range r = Main.instance.wsMTEP.Range[r1, r2];
            r = r.Resize[r.Rows.Count, 2];
            object[,] tep = (object[,])r.Value;
            var dict = Enumerable.Range(1, tep.GetLength(0)).ToDictionary(i => tep[i, 1], i => tep[i, 2]);
            var dict1 = dict.Where(n => n.Value != null).ToDictionary(kv => kv.Key, kv => kv.Value);
            foreach (var item in dict1.Keys)
            {

                txtForCDU += item + "\t" + Math.Round((double)dict[item], 3) + "\n";
            }
            
            using (StreamWriter sw = new StreamWriter(Main.dir + "\\Рапорт.txt"))
            {
                sw.Write(txtForCDU);
            }
        }
    }
}

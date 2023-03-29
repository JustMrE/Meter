using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;

namespace Meter.Forms
{
    public partial class AddPlan : Form
    {
        public ReferenceObject referenceObject;
        public ChildObject childObject;
        public AddPlan(ReferenceObject ro)
        {
            childObject = null;
            referenceObject= ro;
            InitializeComponent();
        }

        public AddPlan(ChildObject co)
        {
            referenceObject= null;
            childObject = co;
            InitializeComponent();
            this.Text = "Код для ТЭП";
            this.tbCod.PlaceholderText = "Введите код для ТЭП";
            this.label3.Text = "Введите код для ТЭП:";
        }

        private void tbCod_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if (!Char.IsDigit(c) && c != 8)
            {
                e.Handled = true;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            if (!string.IsNullOrEmpty(tbCod.Text))
            {
                int cod = Convert.ToInt32(tbCod.Text);
                if (referenceObject != null)
                {
                    if (referenceObject.codPlan == null)
                    {
                        AddNewCodPlan(cod);
                    }
                    else
                    {
                        ChangeCodPlan(cod);
                    }
                }
                else if (childObject != null)
                {
                    if (/*cod < 1301 || cod > 1499*/Main.instance.wsMTEP.Range["A:A"].Find(cod) == null) 
                    {
                        MessageBox.Show("На листе макетТЭПн отсутствует код " + cod + "!\nДобавте код на лист макетТЭПн");
                        return;
                    }
                    else
                    {
                        if (childObject.codTEP == null)
                        {
                            AddNewCodTEP(cod);
                        }
                        else
                        {
                            ChangeCodTEP(cod);
                        }
                    }
                }
            }
            else
            {
                Close();
            }
        }

        private void AddNewCodPlan(int? cod)
        {
            if (cod != 0)
            {
                ReferenceObject ro = Main.instance.references.references.Values.Where(n => n.codPlan == cod).FirstOrDefault();

                if (ro == null)
                {
                    
                    if (!referenceObject.DB.HasItem("план"))
                    {
                        Main.instance.StopAll();
                        referenceObject.AddPlans(false);
                        Main.instance.ResumeAll();
                    }
                    referenceObject.codPlan = cod;
                    Close();
                }
                else
                {
                    if (/*ReplaceDialogue(cod, ro)*/MessageBox.Show(text: "Код " + cod + " уже используется для \"" + ro._name + "\". \nХотите заменить?", caption: "", buttons: MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        ro.codPlan = null;
                        if (!referenceObject.DB.HasItem("план"))
                        {
                            Main.instance.StopAll();
                            referenceObject.AddPlans(false);
                            Main.instance.ResumeAll();
                        }
                        referenceObject.codPlan = cod;
                        Close();
                    }
                }
            }
        }

        private void ChangeCodPlan(int? cod)
        {
            if (cod != 0)
            {
                ReferenceObject ro = Main.instance.references.references.Values.Where(n => n.codPlan == cod).FirstOrDefault();

                if (ro != null && ro != referenceObject)
                {
                    if (MessageBox.Show(text: "Код " + cod + " уже используется для \"" + ro._name + "\". \nХотите заменить?", caption: "", buttons: MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        ro.codPlan = null;
                        referenceObject.codPlan = cod;
                    }
                    Close();
                }
                else
                {
                    referenceObject.codPlan = cod;
                    Close();
                }
            }
            else
            {
                DialogResult result = MessageBox.Show(text: "Хотите удалить код для \"" + referenceObject._name + "\"?", caption: "", buttons: MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    referenceObject.codPlan = null;
                    Close();
                }
            }
        }

        private void AddNewCodTEP(int? cod)
        {
            if (cod != 0)
            {
                ChildObject co = Main.instance.references.references.Values.SelectMany(n => n.PS.childs.Values).Where(m => m.codTEP == cod).FirstOrDefault();

                if (co == null)
                {
                    childObject.codTEP = cod;
                    Main.instance.wsMTEP.Range["A:A"].Find(cod).Interior.Color = Color.GreenYellow;
                    Main.instance.wsMTEP.Range["A:A"].Find(cod).Offset[0, 2].Value = childObject.GetFirstParent._name + " " + childObject._name;
                    Close();
                }
                else
                {
                    if (MessageBox.Show(text: "Код " + cod + " уже используется для \"" + co.GetFirstParent._name + " " + co._name + "\". \nХотите заменить?", caption: "", buttons: MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        co.codTEP = null;
                        childObject.codTEP = cod;
                        Main.instance.wsMTEP.Range["A:A"].Find(cod).Interior.Color = Color.GreenYellow;
                        Main.instance.wsMTEP.Range["A:A"].Find(cod).Offset[0, 2].Value = childObject.GetFirstParent._name + " " + childObject._name;
                        Close();
                    }
                }
            }
        }
        private void ChangeCodTEP(int? cod)
        {
            if (cod != 0)
            {
                ChildObject co = Main.instance.references.references.Values.SelectMany(n => n.PS.childs.Values).Where(m => m.codTEP == cod).FirstOrDefault();

                if (co == null)
                {
                    Main.instance.wsMTEP.Range["A:A"].Find(childObject.codTEP).Interior.ColorIndex = 0;
                    Main.instance.wsMTEP.Range["A:A"].Find(childObject.codTEP).Offset[0, 2].Value = "";
                    childObject.codTEP = cod;
                    Main.instance.wsMTEP.Range["A:A"].Find(cod).Interior.Color = Color.GreenYellow;
                    Main.instance.wsMTEP.Range["A:A"].Find(cod).Offset[0, 2].Value = childObject.GetFirstParent._name + " " + childObject._name;
                    Close();
                }
                else
                {
                    if (MessageBox.Show(text: "Код " + cod + " уже используется для \"" + co.GetFirstParent._name + " " + co._name + "\". \nХотите заменить?", caption: "", buttons: MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        Main.instance.wsMTEP.Range["A:A"].Find(childObject.codTEP).Interior.ColorIndex = 0;
                        Main.instance.wsMTEP.Range["A:A"].Find(childObject.codTEP).Offset[0, 2].Value = "";
                        co.codTEP = null;
                        childObject.codTEP = cod;
                        Main.instance.wsMTEP.Range["A:A"].Find(cod).Interior.Color = Color.GreenYellow;
                        Main.instance.wsMTEP.Range["A:A"].Find(cod).Offset[0, 2].Value = childObject.GetFirstParent._name + " " + childObject._name;
                        Close();
                    }
                }
            }
            else
            {
                DialogResult result = MessageBox.Show(text: "Хотите удалить код для \"" + childObject.GetFirstParent._name + " " + childObject._name + "\"?", caption: "", buttons: MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    Main.instance.wsMTEP.Range["A:A"].Find(childObject.codTEP).Interior.ColorIndex = 0;
                    Main.instance.wsMTEP.Range["A:A"].Find(childObject.codTEP).Offset[0, 2].Value = "";
                    childObject.codTEP = null;
                    Close();
                }
            }
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            this.Close();
        }

        private void AddPlan_Shown(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this);
            if (referenceObject != null)
            {
                if (referenceObject.codPlan != null)
                {
                    label1.Visible = true;
                    label2.Visible = true;
                    label2.Text = referenceObject.codPlan.ToString();
                }
            }
            else if (childObject != null)
            {
                if (childObject.codTEP != null)
                {
                    label1.Visible = true;
                    label2.Visible = true;
                    label2.Text = childObject.codTEP.ToString();
                }
            }
        }
    }
}

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
        public AddPlan(ReferenceObject ro)
        {
            referenceObject= ro;
            InitializeComponent();
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

                if (referenceObject.codPlan == null)
                {
                    AddNewCod(cod);
                }
                else
                {
                    ChangeCod(cod);
                }
            }
            else
            {
                Close();
            }
            
        }

        private void AddNewCod(int? cod)
        {
            if (cod != 0)
            {
                ReferenceObject ro = Main.instance.references.references.Values.Where(n => n.codPlan == cod).FirstOrDefault();

                if (ro == null)
                {
                    
                    if (!referenceObject.DB.HasItem("план"))
                    {
                        referenceObject.AddPlans();
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
                            referenceObject.AddPlans();
                        }
                        referenceObject.codPlan = cod;
                        Close();
                    }
                }
            }
        }

        private void ChangeCod(int? cod)
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

        private void btnCancel_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            this.Close();
        }

        private void AddPlan_Shown(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this);
            if (referenceObject.codPlan != null)
            {
                label1.Visible = true;
                label2.Visible = true;
                label2.Text = referenceObject.codPlan.ToString();
            }
        }
    }
}

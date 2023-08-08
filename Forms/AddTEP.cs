using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;

namespace Meter.Forms
{
    public partial class AddTEP : Form
    {
        public ChildObject childObject;

        public AddTEP(ChildObject co)
        {
            childObject = co;
            InitializeComponent();
            this.Text = childObject.GetFirstParent._name + " " + childObject._name;//"Код для ТЭП";
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

                if (childObject.codTEP == null)
                {
                    AddNewCodTEP(cod);
                }
                else
                {
                    ChangeCodTEP(cod);
                }
            }
            else
            {
                Close();
            }
        }

        private void AddNewCodTEP(int? cod)
        {
            if (cod != 0)
            {
                ChildObject co = Main.instance.references.references.Values.SelectMany(n => n.PS.childs.Values).Where(m => m.codTEP == cod).FirstOrDefault();

                if (co == null)
                {
                    int column = Convert.ToInt32((double)Main.instance.wsTEPm.Range["A1"].Value) + 1;
                    ((Excel.Range)Main.instance.wsTEPn.Cells[2, column]).Resize[5, 2].Borders.LineStyle = XlLineStyle.xlContinuous;
                    ((Excel.Range)Main.instance.wsTEPm.Cells[2, column]).Resize[5, 2].Borders.LineStyle = XlLineStyle.xlContinuous;

                    ((Excel.Range)Main.instance.wsTEPn.Cells[2, column]).Resize[1, 2].Merge();
                    ((Excel.Range)Main.instance.wsTEPn.Cells[2, column]).Value = childObject.GetFirstParent._name + " " + childObject._name;

                    ((Excel.Range)Main.instance.wsTEPn.Cells[3, column]).Value = "план";
                    ((Excel.Range)Main.instance.wsTEPn.Cells[3, column + 1]).Value = "факт";

                    ((Excel.Range)Main.instance.wsTEPn.Cells[4, column]).FormulaR1C1 = "=@INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) - 1,4,1),TRUE) + 1";
                    ((Excel.Range)Main.instance.wsTEPn.Cells[4, column + 1]).FormulaR1C1 = "=@INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) - 1,4,1),TRUE) + 1";

                    ((Excel.Range)Main.instance.wsTEPn.Cells[5, column]).Value = cod;
                    ((Excel.Range)Main.instance.wsTEPn.Cells[5, column + 1]).Value = 0;

                    ((Excel.Range)Main.instance.wsTEPm.Cells[2, column]).Resize[1, 2].Merge();
                    ((Excel.Range)Main.instance.wsTEPm.Cells[2, column]).Value = childObject.GetFirstParent._name + " " + childObject._name;

                    ((Excel.Range)Main.instance.wsTEPm.Cells[3, column]).Value = "план";
                    ((Excel.Range)Main.instance.wsTEPm.Cells[3, column + 1]).Value = "факт";

                    ((Excel.Range)Main.instance.wsTEPm.Cells[4, column]).FormulaR1C1 = "=@INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) - 1,4,1),TRUE) + 1";
                    ((Excel.Range)Main.instance.wsTEPm.Cells[4, column + 1]).FormulaR1C1 = "=@INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) - 1,4,1),TRUE) + 1";

                    ((Excel.Range)Main.instance.wsTEPm.Cells[5, column]).Value = cod;
                    ((Excel.Range)Main.instance.wsTEPm.Cells[5, column + 1]).Value = 0;

                    childObject.codTEP = cod;

                    Close();
                }
                else
                {
                    if (MessageBox.Show(text: "Код " + cod + " уже используется для \"" + co.GetFirstParent._name + " " + co._name + "\". \nХотите заменить?", caption: "", buttons: MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {

                        int column = Convert.ToInt32((double)Main.instance.wsTEPm.Range["5:5"].Find(What: co.codTEP, LookAt: XlLookAt.xlWhole).Column);

                        ((Excel.Range)Main.instance.wsTEPn.Cells[2, column]).Value = childObject.GetFirstParent._name + " " + childObject._name;
                        ((Excel.Range)Main.instance.wsTEPm.Cells[2, column]).Value = childObject.GetFirstParent._name + " " + childObject._name;

                        co.codTEP = null;
                        childObject.codTEP = cod;
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
                    int column = Convert.ToInt32((double)Main.instance.wsTEPm.Range["A1"].Value) + 1;

                    ((Excel.Range)Main.instance.wsTEPn.Cells[2, column]).Resize[5, 2].Borders.LineStyle = XlLineStyle.xlContinuous;
                    ((Excel.Range)Main.instance.wsTEPm.Cells[2, column]).Resize[5, 2].Borders.LineStyle = XlLineStyle.xlContinuous;

                    ((Excel.Range)Main.instance.wsTEPn.Cells[2, column]).Resize[1, 2].Merge();
                    ((Excel.Range)Main.instance.wsTEPn.Cells[2, column]).Value = childObject.GetFirstParent._name + " " + childObject._name;

                    ((Excel.Range)Main.instance.wsTEPn.Cells[3, column]).Value = "план";
                    ((Excel.Range)Main.instance.wsTEPn.Cells[3, column + 1]).Value = "факт";

                    ((Excel.Range)Main.instance.wsTEPn.Cells[4, column]).FormulaR1C1 = "=@INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) - 1,4,1),TRUE) + 1";
                    ((Excel.Range)Main.instance.wsTEPn.Cells[4, column + 1]).FormulaR1C1 = "=@INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) - 1,4,1),TRUE) + 1";

                    ((Excel.Range)Main.instance.wsTEPn.Cells[5, column]).Value = cod;
                    ((Excel.Range)Main.instance.wsTEPn.Cells[5, column + 1]).Value = 0;

                    ((Excel.Range)Main.instance.wsTEPm.Cells[2, column]).Resize[1, 2].Merge();
                    ((Excel.Range)Main.instance.wsTEPm.Cells[2, column]).Value = childObject.GetFirstParent._name + " " + childObject._name;

                    ((Excel.Range)Main.instance.wsTEPm.Cells[3, column]).Value = "план";
                    ((Excel.Range)Main.instance.wsTEPm.Cells[3, column + 1]).Value = "факт";

                    ((Excel.Range)Main.instance.wsTEPm.Cells[4, column]).FormulaR1C1 = "=@INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) - 1,4,1),TRUE) + 1";
                    ((Excel.Range)Main.instance.wsTEPm.Cells[4, column + 1]).FormulaR1C1 = "=@INDIRECT(ADDRESS(ROW(RC),COLUMN(RC) - 1,4,1),TRUE) + 1";

                    ((Excel.Range)Main.instance.wsTEPm.Cells[5, column]).Value = cod;
                    ((Excel.Range)Main.instance.wsTEPm.Cells[5, column + 1]).Value = 0;

                    childObject.codTEP = cod;

                    Close();
                }
                else
                {
                    if (MessageBox.Show(text: "Код " + cod + " уже используется для \"" + co.GetFirstParent._name + " " + co._name + "\". \nХотите заменить?", caption: "", buttons: MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {

                        int column = Convert.ToInt32((double)Main.instance.wsTEPm.Range["5:5"].Find(What: co.codTEP, LookAt: XlLookAt.xlWhole).Column);

                        ((Excel.Range)Main.instance.wsTEPn.Cells[2, column]).Value = childObject.GetFirstParent._name + " " + childObject._name;
                        ((Excel.Range)Main.instance.wsTEPm.Cells[2, column]).Value = childObject.GetFirstParent._name + " " + childObject._name;

                        co.codTEP = null;
                        childObject.codTEP = cod;
                        Close();
                    }
                }
            }
            else
            {
                DialogResult result = MessageBox.Show(text: "Хотите удалить код для \"" + childObject.GetFirstParent._name + " " + childObject._name + "\"?", caption: "", buttons: MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    string adr1 = Main.instance.wsTEPm.Range["5:5"].Find(What: childObject.codTEP, LookAt: XlLookAt.xlWhole).Address[false, false];
                    string adr2 = Main.instance.wsTEPm.Range["5:5"].Find(What: childObject.codTEP, LookAt: XlLookAt.xlWhole).Offset[0, 1].Address[false, false];
                    string adr = Regex.Replace(adr1, @"[^A-Z]+", String.Empty) + ":" + Regex.Replace(adr2, @"[^A-Z]+", String.Empty);

                    Main.instance.StopAll();
                    Main.instance.wsTEPn.Range[adr].Delete(XlDeleteShiftDirection.xlShiftToLeft);
                    Main.instance.wsTEPm.Range[adr].Delete(XlDeleteShiftDirection.xlShiftToLeft);
                    Main.instance.ResumeAll();

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

            if (childObject.codTEP != null)
            {
                label1.Visible = true;
                label2.Visible = true;
                label2.Text = childObject.codTEP.ToString();
            }

        }
    }
}

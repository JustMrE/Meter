using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Meter.Forms
{
    public partial class SCH : Form
    {
        ReferenceObject referenceObject;
        int day;
        double prev, next, coef, sum;
        CultureInfo culture;

        public SCH(ReferenceObject referenceObject)
        {
            culture = CultureInfo.InvariantCulture;
            this.referenceObject = referenceObject;
            day = (int)referenceObject.ActiveDay();
            InitializeComponent();
        }

        private void SCH_Shown(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this);
            var coefVal = referenceObject.DB.childs[RangeReferences.ActiveL1].childs["по счетчику"].RangeByDay(0).Value;
            if (coefVal == null || !double.TryParse(coefVal.ToString(), out coef))
            {
                coef = 0;
            }
            lCoef.Text = coef.ToString();
            
            var prevVal = referenceObject.DB.childs[RangeReferences.ActiveL1].childs["счетчик"].RangeByDay(day - 1).Value;
            if (prevVal == null || !double.TryParse(prevVal.ToString(), out prev))
            {
                prev = 0;
            }
            tbPrev.Text = prev.ToString();

            var nextVal = referenceObject.DB.childs[RangeReferences.ActiveL1].childs["счетчик"].RangeByDay(day).Value;
            if (nextVal == null || !double.TryParse(nextVal.ToString(), out next))
            {
                next = 0;
            }
            tbNext.Text = next.ToString();

            sum = Math.Round((next - prev) * coef, 3);
            lSum.Text = sum.ToString();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            if (sum >= 0)
            {
                referenceObject.WriteToDB(RangeReferences.ActiveL1, "счетчик", (int)day - 1, tbPrev.Text);
                referenceObject.WriteToDB(RangeReferences.ActiveL1, "счетчик", (int)day, tbNext.Text);

                GlobalMethods.ToLog("Измены показания счетчика {" + referenceObject._name + "} " + RangeReferences.ActiveL1 + " за " + ((int)day - 1) + " число на " + tbPrev.Text);
                GlobalMethods.ToLog("Измены показания счетчика {" + referenceObject._name + "} " + RangeReferences.ActiveL1 + " за " + ((int)day) + " число на " + tbNext.Text);

                Close();
            }
            else
            {
                MessageBox.Show("Текущие показания не могут быть меньше чем предыдущие!");
            }
        }

        private void tbPrev_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            string tbVal = tbPrev.Text;
            Check(e, c, tbVal);
        }

        private void tbNext_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            string tbVal = tbNext.Text;
            Check(e, c, tbVal);
        }

        private void tbPrev_TextChanged(object sender, EventArgs e)
        {
            string tbVal = tbPrev.Text;
            GlobalMethods.ToLog(this, sender, tbVal);
            if (string.IsNullOrEmpty(tbVal) || tbVal == "-") return;

            if (!double.TryParse(tbVal, out prev))
            {
                prev = 0;
            }
            sum = Math.Round((next - prev) * coef, 3);
            lSum.Text = sum.ToString();
        }

        private void tbNext_TextChanged(object sender, EventArgs e)
        {
            string tbVal = tbNext.Text;
            GlobalMethods.ToLog(this, sender, tbVal);
            if (string.IsNullOrEmpty(tbVal) || tbVal == "-") return;

            if (!double.TryParse(tbVal, out next))
            {
                next = 0;
            }
            sum = Math.Round((next - prev) * coef, 3);
            lSum.Text = sum.ToString();
        }

        private void Check(KeyPressEventArgs e, char c, string tbVal)
        {
            if (Char.IsDigit(c))
            {

            }
            else if (c == 8)
            {

            }
            else if (c == 44)
            {
                if (tbVal.Contains(",") || string.IsNullOrEmpty(tbVal) || !Char.IsDigit(tbVal.Last()))
                {
                    e.Handled = true;
                }
            }
            else
            {
                e.Handled = true;
            }
        }
    }
}

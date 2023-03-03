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
    public partial class Correct : Form
    {
        ReferenceObject referenceObject;
        int? day;
        double val, cor, sum;
        bool changed;
        string nameL2;

        public Correct(ReferenceObject referenceObject)
        {
            changed = false;
            this.referenceObject = referenceObject;
            day = referenceObject.ActiveDay();
            InitializeComponent();
        }

        private void Correct_Shown(object sender, EventArgs e)
        {
            changed = true;

            var oldStringVal = referenceObject.DB.childs[RangeReferences.ActiveL1].childs["основное"].RangeByDay((int)day).Value;
            double oldSumVal;
            if (oldStringVal == null || !double.TryParse(oldStringVal.ToString(), out oldSumVal))
            {
                oldSumVal = 0;
            }

            if (RangeReferences.ActiveL1 == "план")
            {
                nameL2 = "корректировка";
            }
            else
            {
                nameL2 = "корректировка факт";
            }

            var corVal = referenceObject.DB.childs[RangeReferences.ActiveL1].childs[nameL2].RangeByDay((int)day).Value;
            double corrVal;
            if (corVal == null || !double.TryParse(corVal.ToString(), out corrVal))
            {
                corrVal = 0;
            }
            val = Math.Round(oldSumVal - corrVal, 3);
            cor = Math.Round(corrVal, 3);
            sum = Math.Round(oldSumVal, 3);

            label4.Text = val.ToString();
            tbCorr.Text = cor.ToString();
            tbSum.Text = sum.ToString();
            
            changed = false;
        }

        private void tbSum_TextChanged(object sender, EventArgs e)
        {
            string tbVal = tbSum.Text;
            if (!changed) 
            {
                if (string.IsNullOrEmpty(tbSum.Text) || tbVal == "-")
                {
                    changed = true;
                    tbCorr.Text = "";
                    changed = false;
                    return;
                }
                changed = true;

                if (!double.TryParse(tbVal, out sum))
                {
                    sum = 0;
                }
                double result = Math.Round(sum - val, 3);
                tbCorr.Text = result.ToString();
                changed = false;
            }
        }

        private void tbCorr_TextChanged(object sender, EventArgs e)
        {
            string tbVal = tbCorr.Text;
            if (!changed)
            {
                if (string.IsNullOrEmpty(tbVal) || tbVal == "-")
                {
                    changed = true;
                    tbSum.Text = val.ToString();
                    changed = false;
                    return;
                }
                changed = true;

                if (!double.TryParse(tbVal, out cor))
                {
                    cor = 0;
                }
                double result = Math.Round(val + cor, 3);
                tbSum.Text = result.ToString();
                changed = false;
            }
            
        }

        private void tbCorr_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            string tbVal = tbCorr.Text;
            Check(e, c, tbVal);
        }

        private void tbSum_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            string tbVal = tbSum.Text;
            Check(e, c, tbVal);
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
            else if (c == 45)
            {
                if (tbVal.Contains("-") || !string.IsNullOrEmpty(tbVal))
                {
                    e.Handled = true;
                }
            }
            else
            {
                e.Handled = true;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            referenceObject.WriteToDB(RangeReferences.ActiveL1, nameL2, (int)day, tbCorr.Text);
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}

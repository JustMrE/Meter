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
    public partial class ChangeCoef : Form
    {
        ReferenceObject referenceObject;

        public ChangeCoef(ReferenceObject referenceObject)
        {
            this.referenceObject = referenceObject;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            string newValue = textBox1.Text;
            if (!string.IsNullOrEmpty(newValue))
            {
                referenceObject.meterCoef = newValue;
            }
            else
            {
                referenceObject.meterCoef = null;
            }
            
            referenceObject.UpdateMeterCoef();
            Close();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            string tbVal = textBox1.Text;
            Check(e, c, tbVal);
        }

        private void ChangeCoef_Shown(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this);
            string? oldValue = referenceObject.meterCoef;
            if (!string.IsNullOrEmpty(oldValue))
            {
                label3.Text= oldValue;
            }
            else
            {
                label2.Visible = false;
                label3.Visible = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            Close();
        }

        private void Check(KeyPressEventArgs e, char c, string tbVal)
        {
            if (Char.IsDigit(c))
            {

            }
            else if (c == 8)
            {

            }
            else if (c == 44 )
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

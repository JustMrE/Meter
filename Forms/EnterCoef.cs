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
    public partial class EnterCoef : Form
    {
        public string oldVal, newVal;
        Control button;

        public EnterCoef(Control b)
        {
            button = b;
            InitializeComponent();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            string newValue = textBox1.Text;
            if (string.IsNullOrEmpty(newValue))
            {
                //referenceObject.meterCoef = newValue;
                MessageBox.Show("Значение не может быть пустым!");
            }
            else
            {
                //referenceObject.meterCoef = null;
                newVal = textBox1.Text;
                Close();
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            string tbVal = textBox1.Text;
            Check(e, c, tbVal);
        }

        private void ChangeCoef_Shown(object sender, EventArgs e)
        {
            oldVal = button.Text;
            textBox1.Text = oldVal;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
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

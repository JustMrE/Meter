using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Main = Meter.MyApplicationContext;

namespace Meter.Forms
{
    public partial class NewMenuBase : Form
    {
        public NewMenuBase()
        {
            InitializeComponent();
        }

        protected virtual void button1_Click(object sender, EventArgs e)
        {

        }

        protected virtual void NewMenuBase_Shown(object sender, EventArgs e)
        {
            SetWindowLong(this.Handle, GWL_EXSTYLE, WS_EX_TOOLWINDOW);
        }

        protected virtual void NewMenuBase_Load(object sender, EventArgs e)
        {
            formHwnd = this.Handle;
            SetParent(formHwnd, Main.instance.xlAppHwnd);
            
        }

        protected virtual void NewMenuBase_Activated(object sender, EventArgs e)
        {
            TempForm.instance.CalculateFormsPositions();
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace Meter.Forms
{
    public partial class NewMenu : Meter.Forms.NewMenuBase
    {
        public NewMenu() : base()
        {
            InitializeComponent();
        }

        protected override void button1_Click(object sender, EventArgs e)
        {
            base.button1_Click(sender, e);
            this.Hide();
            TempForm.instance.menues[1].Show();
        }

        protected override void NewMenuBase_Shown(object sender, EventArgs e)
        {
            base.NewMenuBase_Shown(sender, e);
            TempForm.instance.menuForm = this;
        }
    }
}

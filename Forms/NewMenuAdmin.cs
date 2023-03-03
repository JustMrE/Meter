using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Meter.Forms
{
    public partial class NewMenuAdmin : Meter.Forms.NewMenuBase
    {
        public NewMenuAdmin() : base()
        {
            InitializeComponent();
        }

        protected override void button1_Click(object sender, EventArgs e)
        {
            base.button1_Click(sender, e);
            TempForm.instance.menues[0].Show();
            this.Hide();
        }

        protected override void NewMenuBase_Shown(object sender, EventArgs e)
        {
            base.NewMenuBase_Shown(sender, e);
        }

        protected override void NewMenuBase_Load(object sender, EventArgs e)
        {
            base.NewMenuBase_Load(sender, e);
        }

        protected override void NewMenuBase_Activated(object sender, EventArgs e)
        {
            TempForm.instance.menuForm = this;
            base.NewMenuBase_Activated(sender, e);
        }
    }
}

using System.ComponentModel;

namespace Meter
{
    public abstract class MyFormBase : Form
    {
        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
            GlobalMethods.ToLog(this, true);
        }

        protected virtual void btnCancel_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            this.Close();
        }
    }
}
using System.Runtime.InteropServices;
using Main = Meter.MyApplicationContext;

namespace Meter.Forms
{
    public partial class SplashScreen : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        protected static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
        

        public SplashScreen()
        {
            InitializeComponent();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            
        }

        public void UpdateText(string txt)
        {
            Label1.Text = txt;
        }

        public void UpdateLabel(string txt)
        {
            Label0.Text = txt;
        }
    }
}

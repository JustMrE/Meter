using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Meter.Forms
{
    partial class NewMenuBase
    {
        [DllImport("user32.dll", SetLastError = true)]
        protected static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        protected static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOOLWINDOW = 0x80;
        public IntPtr formHwnd;

        //public NewMenuBase()
        //{
        //    InitializeComponent();
        //    SetWindowLong(this.Handle, GWL_EXSTYLE, WS_EX_TOOLWINDOW);
        //}

        public virtual void SetRects(int left, int top, int width, int height)
        {
            Action<int, int, int, int> action = SetPosition;
            if (InvokeRequired)
            {
                Invoke(action, left, top, width, height);
            }
            else
            {
                action(left, top, width, height);
            }
        }
        private void SetPosition(int left, int top, int width, int height)
        {
            Location = new System.Drawing.Point(left, top);
            Width = width;
            Height = height;
        }
    }
}

namespace Meter.Forms
{
    public partial class DBClearing : Form
    {
        public DBClearing()
        {
            InitializeComponent();
            Label1.Text = "";
        }

        public void UpdateText(string txt)
        {
            Label1.Text = txt;
        }
    }
}
using Main = Meter.MyApplicationContext;

namespace Meter.Forms
{
    partial class AddSubject : Form
    {
        string? nameL0, nameL1, nameL2;
        public AddSubject()
        {
            InitializeComponent();
            ComboBox11.DataSource = Main.instance.heads.heads.Keys.ToList();
            ComboBox11.Text = "";
        }

        public void ComboBox11_TextChanged(object sender, EventArgs e)
        {
            nameL0 = ComboBox11.Text;
            if (!string.IsNullOrEmpty(nameL0) && Main.instance.heads.heads.ContainsKey(nameL0))
            {
                ComboBox12.DataSource = Main.instance.heads.heads[nameL0].childs.Keys.ToList();
                ComboBox12.Text = "";
                ComboBox12.Visible = true;
            }
            else
            {
                nameL1 = null;
                nameL2 = null;
                ComboBox12.DataSource = null;
                ComboBox13.DataSource = null;
                ComboBox12.Visible = false;
                ComboBox13.Visible = false;
            }
        }

        public void ComboBox12_TextChanged(object sender, EventArgs e)
        {
            nameL1 = ComboBox12.Text;
            if (!string.IsNullOrEmpty(nameL0) && !string.IsNullOrEmpty(nameL1) && Main.instance.heads.heads[nameL0].childs.ContainsKey(nameL1))
            {
                ComboBox13.DataSource = Main.instance.heads.heads[nameL0].childs[nameL1].childs.Keys.ToList();
                ComboBox13.Text = "";
                ComboBox13.Visible = true;
            }
            else
            {
                nameL2 = null;
                ComboBox13.DataSource = null;
                ComboBox13.Visible = false;
            }
        }

        public void ComboBox13_TextChanged(object sender, EventArgs e)
        {
            nameL2 = ComboBox13.Text;
        }

        public void btnOk_Click(object sender, EventArgs e)
        {
            Close();
        }

        public void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
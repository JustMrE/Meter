using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace Meter
{

    partial class SaveFormula : Form
    {
        private string path1;
        private List<ForTags> saveFormula;
        private Dictionary<string, string> formulas = new();

        public SaveFormula(List<ForTags> saveFormula)
        {
            this.saveFormula = saveFormula;
            InitializeComponent();
            path1 = Main.dir + @"\saves";
            if (!Directory.Exists(path1)) Directory.CreateDirectory(path1);
            path1 = Main.dir + @"\saves\formulas";
            if (!Directory.Exists(path1)) Directory.CreateDirectory(path1);
            List<string> filePaths1 = Directory.GetFiles(path1, "*.json").ToList();
            foreach (string path in filePaths1)
            {
                try 
                {
                    string name = path.Replace(path1 + @"\","").Replace(".json", "");
                    ListBox1.Items.Add(name);
                    formulas.Add(name, path);
                }
                catch (Exception) { }
            }
        }

        protected void btnOK_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbName.Text))
            {
                MessageBox.Show("звание не может быть пустым!");
                return;
            }
            string filename = path1 + @"\" + tbName.Text + ".json";
            if(ListBox1.Items.Contains(tbName.Text) || File.Exists(filename))
            {
                if (MessageBox.Show("Такое сохранение уже существует! Хотите заменить?", "", MessageBoxButtons.YesNo) == DialogResult.No) return;
            }
            using (StreamWriter writer = File.CreateText(filename))
            {
                JsonSerializer serializer = new JsonSerializer();
                var jsonString = JsonConvert.SerializeObject(saveFormula);
                writer.Write(jsonString);
            }
            this.DialogResult = DialogResult.OK;
            Close();
        }
        protected void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            Close();
        }

        protected void ListBox1_Click(object sender, EventArgs e)
        {
            if (ListBox1.SelectedItems.Count != 0)
            {
                btnDelete.Enabled = true;
                tbName.Text = ListBox1.SelectedItems[0].ToString();
            }
            else
            {
                btnDelete.Enabled = false;
            }
        }

        protected void btnDelete_Click(object sender, EventArgs e)
        {
            string deleteFile = formulas[ListBox1.SelectedItem.ToString()];
            try
            {
                if (File.Exists(deleteFile))
                {
                    File.Delete(deleteFile);
                }
                formulas.Remove(ListBox1.SelectedItem.ToString());
                ListBox1.Items.Remove(ListBox1.SelectedItem.ToString());
            }
            catch (Exception) {}
        }
    }
}
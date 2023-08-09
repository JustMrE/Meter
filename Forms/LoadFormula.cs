using Main = Meter.MyApplicationContext;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace Meter
{

    partial class LoadFormula : Form
    {
        private List<string> list = new ();
        private Dictionary<string, string> formulas = new();
        public List<ForTags> loadedFormula;
        private string path1;

        public LoadFormula()
        {
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
            if (ListBox1.SelectedItems.Count == 0)
            {
                Close();
            }
            GlobalMethods.ToLog(this, sender);
            OpenFormula();
            this.DialogResult = DialogResult.OK;
            Close();
        }
        protected void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            Close();
        }

        private void OpenFormula()
        {
            string loadFile = formulas[ListBox1.SelectedItems[0].ToString()];

            using (var streamReader = new StreamReader(loadFile))
            {
                var json = streamReader.ReadToEnd();
                loadedFormula = JsonConvert.DeserializeObject<List<ForTags>>(json);
            }
        }

        protected void ListBox1_Click(object sender, EventArgs e)
        {
            if (ListBox1.SelectedItems.Count != 0)
            {
                btnDelete.Enabled = true;
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

        private void tbSearch_TextChanged(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender, tbSearch.Text);
            RegexSearch();
        }

        private void RegexSearch()
        {
            // RegexOptions ro = this.сheckBox1.Checked ? RegexOptions.None : RegexOptions.IgnoreCase;
            RegexOptions ro = RegexOptions.IgnoreCase;
            if (this.tbSearch.Text != "")
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    try
                    {
                        ListBox1.Items.Clear();
                        string search = this.tbSearch.Text;
                        search = search.Replace("*", @".*");

                        var deviceIds = list.AsEnumerable();
                        var matchingIds = deviceIds.Where(id => Regex.IsMatch(id, pattern: search, ro)).ToArray();
                        ListBox1.Items.AddRange(matchingIds.ToArray());
                    }
                    catch (ArgumentException)
                    {
                    this.Invoke((MethodInvoker)(() =>
                    {
                        ListBox1.Items.Clear();
                            ListBox1.Items.AddRange(list.OrderBy(m => m).ToArray());
                    }));
                }

                }));
            }
            else
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    ListBox1.Items.Clear();
                    ListBox1.Items.AddRange(list.OrderBy(m => m).ToArray());
                }));
            }

        }
    }
}
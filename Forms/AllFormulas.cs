using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Text.RegularExpressions;

namespace Meter.Forms
{
    public partial class AllFormulas : Form
    {
        private List<string> list = new List<string>();
        private Dictionary<string, string> idByName = new Dictionary<string, string>();

        public AllFormulas()
        {
            InitializeComponent();
            int i;
            foreach (string id in Main.instance.formulas.formulas.Keys)
            {
                string name = ((ChildObject)RangeReferences.idDictionary[id]).GetFirstParent._name + " " + RangeReferences.idDictionary[id]._name;
                i = listBox0.Items.Add(name);
                list.Add(name);
                idByName.Add(name, id);
            }
        }

        private void OpenFormula(string newID)
        {
            ChildObject co1 = (ChildObject)RangeReferences.idDictionary[newID];
            string newNameL1 = co1._name;
            ChildObject co = co1.GetParent<ChildObject>();
            ReferenceObject ro = co.GetFirstParent;

            Thread t = new Thread(() =>
            {
                FormulaEditor form = new FormulaEditor(ref ro, newNameL1);
                form.FormClosed += (s, args) =>
                {
                    System.Windows.Forms.Application.ExitThread();
                };
                form.Show();
                System.Windows.Forms.Application.Run();
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        private void listBox0_DoubleClick(object? sender, System.EventArgs e)
        {
            string name = listBox0.SelectedItem as string;
            string id = idByName[name];
            OpenFormula(id);
        }

        private void RegexSearch()
        {
            RegexOptions ro = RegexOptions.None; //checkBox1.Checked ? RegexOptions.None : RegexOptions.IgnoreCase;
            if (this.tbSearch.Text != "")
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    try
                    {
                        listBox0.Items.Clear();
                        string search = this.tbSearch.Text;
                        search = search.Replace("*", @".*");

                        var deviceIds = list.AsEnumerable();
                        var matchingIds = deviceIds.Where(id => Regex.IsMatch(id, pattern: search, ro)).ToArray();
                        listBox0.Items.AddRange(matchingIds.ToArray());
                    }
                    catch (ArgumentException)
                    {
                    this.Invoke((MethodInvoker)(() =>
                    {
                        listBox0.Items.Clear();
                            listBox0.Items.AddRange(list.OrderBy(m => m).ToArray());
                    }));
                }

                }));
            }
            else
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    listBox0.Items.Clear();
                    listBox0.Items.AddRange(list.OrderBy(m => m).ToArray());
                }));
            }

        }

        private void tbSearch_TextChanged(object sender, EventArgs e)
        {
            RegexSearch();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}

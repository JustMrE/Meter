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
        private List<ListViewItem> list = new List<ListViewItem>();
        private Dictionary<string, string> idByName = new Dictionary<string, string>();

        public AllFormulas()
        {
            InitializeComponent();
            LoadAllFormulas();
        }

        public AllFormulas(ReferenceObject ro, string activeL1)
        {
            InitializeComponent();
            LoadFormulasFor(ro, activeL1);
        }

        private void LoadFormulasFor(ReferenceObject ro, string activeL1) 
        {
            string i = ro.DB.childs[activeL1].childs["основное"].ID;
            List<string> formulas = Main.instance.formulas.formulas.Where(kv => kv.Value.Any(f => f.ID == i)).Select(kv => kv.Key).ToList();
            foreach (string id in formulas)
            {
                try 
                {
                    string name = ((ChildObject)RangeReferences.idDictionary[id]).GetFirstParent._name + " " + RangeReferences.idDictionary[id]._name;
                    ListViewItem item = new ListViewItem()
                    {
                        Text = name,
                    };
                    if (Main.instance.formulas.formulas[id].Where(f => f.type == ButtonsType.subject && f.ID == null).ToList().Count > 0)
                    {
                        item.BackColor = Color.Orange;
                        item.ToolTipText = "В формуле ошибка! Удаленный субъект...";
                    }
                    listView0.Items.Add(item);
                    list.Add(item);
                    idByName.Add(name, id);
                }
                catch (Exception e)
                {
                    ListViewItem item = new ListViewItem()
                    {
                        Text = id,
                        BackColor = Color.Red,
                        ToolTipText = "Субъект не существует в счетчиках..."
                    };
                    listView0.Items.Add(item);
                    list.Add(item);
                    GlobalMethods.Err(e.Message);
                }
            }
        }
        private void LoadAllFormulas()
        {
            foreach (string id in Main.instance.formulas.formulas.Keys)
            {
                try 
                {
                    string name = ((ChildObject)RangeReferences.idDictionary[id]).GetFirstParent._name + " " + RangeReferences.idDictionary[id]._name;
                    // i = listBox0.Items.Add(name);
                    ListViewItem item = new ListViewItem()
                    {
                        Text = name,
                    };
                    if (Main.instance.formulas.formulas[id].Where(f => f.type == ButtonsType.subject && f.ID == null).ToList().Count > 0)
                    {
                        item.BackColor = Color.Orange;
                        item.ToolTipText = "В формуле ошибка! Удаленный субъект...";
                    }
                    listView0.Items.Add(item);
                    list.Add(item);
                    idByName.Add(name, id);
                }
                catch (Exception e)
                {
                    ListViewItem item = new ListViewItem()
                    {
                        Text = id,
                        BackColor = Color.Red,
                        ToolTipText = "Субъект не существует в счетчиках..."
                    };
                    listView0.Items.Add(item);
                    list.Add(item);
                    GlobalMethods.Err(e.Message);
                }
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

        private void listBox0_DoubleClick(object? sender, EventArgs e)
        {
            ListViewItem item = listView0.SelectedItems[0];
            GlobalMethods.ToLog(this, sender, item.Text);
            if (item.BackColor != Color.Red)
            {
                string name = listView0.SelectedItems[0].Text;
                string id = idByName[name];
                OpenFormula(id);
            }
            else
            {
                DialogResult ans = MessageBox.Show("Ошибка! Отсутствует формула или субъект!\nУдалить эту формулу?","", MessageBoxButtons.YesNo);
                if (ans == DialogResult.Yes)
                {
                    if (File.Exists(Main.dir + @"\current\formulas\" + item.Text + ".json"))
                    {
                        Main.filesToDelete.Add(Main.dir + @"\current\formulas\" + item.Text + ".json");
                        listView0.Items.Remove(item);
                        list.Remove(item);
                        Main.instance.formulas.formulas.Remove(item.Text);
                        MessageBox.Show("Удалено!");
                    }
                    else
                    {
                        MessageBox.Show("Не удалось удалить! Удалите вручную.");
                    }
                }
            }
        }

        private void RegexSearch()
        {
            RegexOptions ro = this.сheckBox1.Checked ? RegexOptions.None : RegexOptions.IgnoreCase;
            if (this.tbSearch.Text != "")
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    try
                    {
                        listView0.Items.Clear();
                        string search = this.tbSearch.Text;
                        search = search.Replace("*", @".*");

                        var deviceIds = list.AsEnumerable();
                        var matchingIds = deviceIds.Where(id => Regex.IsMatch(id.Text, pattern: search, ro)).ToArray();
                        listView0.Items.AddRange(matchingIds.ToArray());
                    }
                    catch (ArgumentException)
                    {
                    this.Invoke((MethodInvoker)(() =>
                    {
                        listView0.Items.Clear();
                            listView0.Items.AddRange(list.OrderBy(m => m.Text).ToArray());
                    }));
                }

                }));
            }
            else
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    listView0.Items.Clear();
                    listView0.Items.AddRange(list.OrderBy(m => m.Text).ToArray());
                }));
            }

        }
        
        private void tbSearch_TextChanged(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender, tbSearch.Text);
            RegexSearch();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            Close();
        }

        private void AllFormulas_Shown(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this);
        }
    }
}

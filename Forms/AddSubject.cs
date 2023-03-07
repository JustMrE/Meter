//using Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;

namespace Meter.Forms
{
    partial class AddSubject : Form
    {
        string? nameL0, nameL1, nameL2;
        public string adr;
        List<string> types = new List<string>();
        public AddSubject()
        {
            InitializeComponent();
            ComboBox11.DataSource = Main.instance.heads.heads.Keys.ToList();
        }

        private void AddSubject_Shown(object sender, EventArgs e)
        {
            ComboBox11.Text = string.Empty;
        }

        public void ComboBox11_TextChanged(object sender, EventArgs e)
        {
            nameL0 = ComboBox11.Text;
            if (!string.IsNullOrEmpty(nameL0))
            {
                ComboBox12.Visible = true;
                if (Main.instance.heads.heads.ContainsKey(nameL0))
                {
                    ComboBox12.DataSource = Main.instance.heads.heads[nameL0].childs.Keys.ToList();
                    ComboBox12.Text = string.Empty;
                    
                }
                else
                {
                    ComboBox12.DataSource = null;
                    ComboBox13.DataSource = null;
                }
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
            if (!string.IsNullOrEmpty(nameL1))
            {
                ComboBox13.Visible = true;
                if (Main.instance.heads.heads[nameL0].childs.ContainsKey(nameL1))
                {
                    ComboBox13.DataSource = Main.instance.heads.heads[nameL0].childs[nameL1].childs.Keys.ToList();
                    ComboBox13.Text = string.Empty;
                }
                else
                {
                    ComboBox13.DataSource = null;
                }
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

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = sender as CheckBox;
            string type = cb.Text;  
            if (cb.Checked == true)
            {
                if (!types.Contains(type)) types.Add(type);
            }
            else
            {
                if (types.Contains(type)) types.Remove(type);
            }
        }

        public void btnOk_Click(object sender, EventArgs e)
        {
            string name = TextBox15.Text;
            if (string.IsNullOrEmpty(nameL0) || string.IsNullOrEmpty(nameL1) || string.IsNullOrEmpty(nameL2))
            {
                MessageBox.Show("¬ведите расположение в таблице!");
                return;
            }
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("¬ведите название субекта!");
                return;
            }
            else if (Main.instance.references.references.ContainsKey(name)) 
            {
                MessageBox.Show("—убъект с таким именем уже существует!");
                return;
            }
            if (types.Count == 0)
            {
                MessageBox.Show("¬ыберите хот€бы один вариант! (прием/отдача/сальдо)");
                return;
            }

            
            if (Main.instance.heads.heads.ContainsKey(nameL0))
            {
                if (Main.instance.heads.heads[nameL0].childs.ContainsKey(nameL1))
                {
                    if (Main.instance.heads.heads[nameL0].childs[nameL1].childs.ContainsKey(nameL2))
                    {
                        adr = Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].LastCell.Offset[1, 1].Address[false, false];
                        CreateNew();
                        UpdateUps();
                        Close();
                    }
                    else
                    {
                        Main.instance.StopAll();
                        Excel.Range r = Main.instance.heads.heads[nameL0].childs[nameL1].LastCell.Offset[-1, 1].Resize[42];
                        string address = Main.instance.heads.heads[nameL0].childs[nameL1].LastCell.Offset[1, 1].Resize[1].Address[false, false];
                        r.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                        r = Main.instance.wsCh.Range[address];
                        r.Value = nameL2;
                        Main.instance.heads.heads[nameL0].childs[nameL1].childs.Add(nameL2, new HeadObject()
                        {
                            WS = Main.instance.wsCh,
                            _name = nameL2,
                            Range = r,
                            _level = Level.level2,
                        });

                        Main.instance.ResumeAll();
                        adr = r.Offset[1].Address[false, false];
                        CreateNew(false);

                        UpdateUps();
                        Close();
                    }
                }
                else
                {
                    Main.instance.StopAll();
                    Excel.Range r = Main.instance.heads.heads[nameL0].LastCell.Offset[-1, 1].Resize[42];
                    string address = Main.instance.heads.heads[nameL0].LastCell.Offset[1, 1].Resize[1].Address[false, false];
                    r.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                    r = Main.instance.wsCh.Range[address];
                    r.Value = nameL1;

                    Main.instance.heads.heads[nameL0].childs.Add(nameL1, new HeadObject()
                    {
                        WS = Main.instance.wsCh,
                        _name = nameL1,
                        Range = r,
                        _level = Level.level1,
                        childs = new Dictionary<string, HeadObject>(),
                    });

                    r = r.Offset[1];
                    r.Value = nameL2;

                    Main.instance.heads.heads[nameL0].childs[nameL1].childs.Add(nameL2, new HeadObject()
                    {
                        WS = Main.instance.wsCh,
                        _name = nameL2,
                        Range = r,
                        _level = Level.level2,
                    });

                    Main.instance.ResumeAll();

                    adr = r.Offset[1].Address[false, false];
                    CreateNew(false);

                    UpdateUps();
                    Close();
                }
            }
            else
            {
                
            }
        }

        private void CreateNew(bool insert = true)
        {
            string name = TextBox15.Text;
            
            Main.instance.references.CreateNew(name, types[0], adr, insert);
            types.RemoveAt(0);
            if (types.Count > 0)
            {
                foreach (string t in types)
                {
                    Main.instance.references.references[name].AddNewDBL1StandartOther(t);
                    Main.instance.references.references[name].AddNewPS(t, "ручное");
                    Main.instance.references.references[name].PS.childs[t].childs["ручное"].ChangeCod();
                }
            }
        }

        public void UpdateUps()
        {

            if (Main.instance.heads.heads[nameL0].LastCell.Column < Main.instance.wsCh.Range[adr].Column)
            {
                Main.instance.heads.heads[nameL0].Resize(1, false);
            }
            if (Main.instance.heads.heads[nameL0].childs[nameL1].LastCell.Column < Main.instance.wsCh.Range[adr].Column)
            {
                Main.instance.heads.heads[nameL0].childs[nameL1].Resize(1, false);
            }
            if (Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].LastCell.Column < Main.instance.wsCh.Range[adr].Column)
            {
                Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].Resize(1, false);
            }
            
            Main.instance.heads.heads[nameL0].UpdateColors();
            Main.instance.heads.heads[nameL0].UpdateBorders();
            
            Main.instance.heads.heads[nameL0].childs[nameL1].UpdateColors();
            Main.instance.heads.heads[nameL0].childs[nameL1].UpdateBorders();
            
            Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].UpdateColors();
            Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].UpdateBorders();
        }

        public void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
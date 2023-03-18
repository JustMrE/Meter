//using Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;

namespace Meter.Forms
{
    partial class AddSubject : Form
    {
        string? nameL0, nameL1, nameL2;
        string name;
        public string adr;
        List<string> types = new List<string>();
        public AddSubject()
        {
            InitializeComponent();
            if (Main.instance.heads.heads != null)
            {
                ComboBox11.DataSource = Main.instance.heads.heads.Keys.ToList();
            }
        }

        private void AddSubject_Shown(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this);
            ComboBox11.Text = string.Empty;
        }

        public void ComboBox11_TextChanged(object sender, EventArgs e)
        {
            nameL0 = ComboBox11.Text;
            GlobalMethods.ToLog(this, sender, nameL0);
            if (!string.IsNullOrEmpty(nameL0))
            {
                ComboBox12.Visible = true;
                if (Main.instance.heads.heads.ContainsKey(nameL0))
                {
                    if (Main.instance.heads.heads[nameL0].childs != null)
                    {
                        ComboBox12.DataSource = Main.instance.heads.heads[nameL0].childs.Keys.ToList();
                    }
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
            GlobalMethods.ToLog(this, sender, nameL1);
            if (!string.IsNullOrEmpty(nameL1))
            {
                ComboBox13.Visible = true;
                if (Main.instance.heads.heads.ContainsKey(nameL0) && Main.instance.heads.heads[nameL0].childs.ContainsKey(nameL1))
                {
                    if (Main.instance.heads.heads[nameL0].childs[nameL1].childs != null)
                    {
                        ComboBox13.DataSource = Main.instance.heads.heads[nameL0].childs[nameL1].childs.Keys.ToList();
                    }
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
            GlobalMethods.ToLog(this, sender, nameL2);
        }

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = sender as CheckBox;
            GlobalMethods.ToLog(this, sender, cb.Checked);
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
            GlobalMethods.ToLog(this, sender);
            name = TextBox15.Text + " " + ComboBox13.Text;
            if (string.IsNullOrEmpty(nameL0) || string.IsNullOrEmpty(nameL1) || string.IsNullOrEmpty(nameL2))
            {
                MessageBox.Show("Введите расположение субъекта в таблиуе!");
                return;
            }
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("Введите название субъекта!");
                return;
            }
            else if (Main.instance.references.references.ContainsKey(name)) 
            {
                MessageBox.Show("Субъект с таким название уже существует!");
                return;
            }
            if (types.Count == 0)
            {
                MessageBox.Show("Выберите хотябы один вариант! (прием/отдача/сальдо)");
                return;
            }

            
            if (Main.instance.heads.heads != null && Main.instance.heads.heads.ContainsKey(nameL0))
            {
                if (Main.instance.heads.heads[nameL0].childs != null && Main.instance.heads.heads[nameL0].childs.ContainsKey(nameL1))
                {
                    if (Main.instance.heads.heads[nameL0].childs[nameL1].childs != null && Main.instance.heads.heads[nameL0].childs[nameL1].childs.ContainsKey(nameL2))
                    {
                        Main.instance.StopAll();
                        adr = Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].LastCell.Offset[1, 1].Address[false, false];
                        CreateNew();
                        UpdateUps(false);
                        Main.instance.ResumeAll();
                        Close();
                    }
                    else if (Main.instance.heads.heads[nameL0].childs[nameL1].childs == null)
                    {
                        Main.instance.heads.heads[nameL0].childs[nameL1].childs = new Dictionary<string, HeadObject> ();
                        Main.instance.StopAll();
                        Excel.Range r = Main.instance.heads.heads[nameL0].childs[nameL1].LastCell.Offset[1, 0].Resize[1];
                        //string address = Main.instance.heads.heads[nameL0].childs[nameL1].LastCell.Offset[1, 0].Resize[1].Address[false, false];
                        //r.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                        //r = Main.instance.wsCh.Range[address];
                        r.Value = nameL2;
                        Main.instance.heads.heads[nameL0].childs[nameL1].childs.Add(nameL2, new HeadObject()
                        {
                            WS = Main.instance.wsCh,
                            _name = nameL2,
                            Range = r,
                            _level = Level.level2,
                        });
                        adr = r.Offset[1].Address[false, false];
                        CreateNew(false);
                        UpdateUps(false);
                        Main.instance.ResumeAll();
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
                        adr = r.Offset[1].Address[false, false];
                        CreateNew(false);
                        UpdateUps(false);
                        Main.instance.ResumeAll();
                        Close();
                    }
                }
                else if (Main.instance.heads.heads[nameL0].childs == null)
                {
                    Main.instance.heads.heads[nameL0].childs = new Dictionary<string, HeadObject>();
                    Main.instance.StopAll();
                    Excel.Range r = Main.instance.heads.heads[nameL0].LastCell.Offset[1, 0].Resize[1];
                    //string address = Main.instance.heads.heads[nameL0].LastCell.Offset[1, 1].Resize[1].Address[false, false];
                    //r.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                    //r = Main.instance.wsCh.Range[address];
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
                    adr = r.Offset[1].Address[false, false];
                    CreateNew(false);
                    UpdateUps(false);
                    Main.instance.ResumeAll();
                    Close();
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
                    adr = r.Offset[1].Address[false, false];
                    CreateNew(false);
                    UpdateUps(false);
                    Main.instance.ResumeAll();
                    Close();
                }
            }
            else if (Main.instance.heads.heads == null)
            {
                Main.instance.heads.heads = new Dictionary<string, HeadObject>();
                int row = 8;
                Main.instance.StopAll();
                Excel.Range r = (Excel.Range)Main.instance.wsCh.Cells[row, 2];
                r.Value = nameL0;
                Main.instance.heads.heads.Add(nameL0, new HeadObject()
                {
                    WS = Main.instance.wsCh,
                    _name = nameL0,
                    Range = r,
                    _level = Level.level0,
                    childs = new Dictionary<string, HeadObject>(),
                });
                r = r.Offset[1];
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
                adr = r.Offset[1].Address[false, false];
                CreateNew(false);
                UpdateUps(false);
                Main.instance.ResumeAll();
                Close();
            }
            else
            {
                int row = Main.instance.heads.heads.Values.Select(n => n.Range.Row).Max();
                row = row + 43;
                Main.instance.StopAll();
                Excel.Range r = (Excel.Range)Main.instance.wsCh.Cells[row, 2];
                r.Value = nameL0;
                Main.instance.heads.heads.Add(nameL0, new HeadObject()
                {
                    WS = Main.instance.wsCh,
                    _name = nameL0,
                    Range = r,
                    _level = Level.level0,
                    childs = new Dictionary<string, HeadObject>(),
                });
                r = r.Offset[1];
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
                adr = r.Offset[1].Address[false, false];
                CreateNew(false);
                UpdateUps(false);
                Main.instance.ResumeAll();
                Close();
            }
        }

        private void CreateNew(bool insert = true)
        {            
            Main.instance.references.CreateNew(name, types[0], adr, insert, false);
            types.RemoveAt(0);
            if (types.Count > 0)
            {
                foreach (string t in types)
                {
                    Main.instance.references.references[name].AddNewDBL1StandartOther(t, false);
                    Main.instance.references.references[name].AddNewPS(t, "ручное", false);
                    Main.instance.references.references[name].PS.childs[t].childs["ручное"].ChangeCod();
                }
            }
        }

        public void UpdateUps(bool stopall = true)
        {

            if (Main.instance.heads.heads[nameL0].LastCell.Column < Main.instance.references.references[name].PS.LastColumn.Column)
            {
                int resizeValue = Main.instance.references.references[name].PS.LastColumn.Column - Main.instance.heads.heads[nameL0].LastCell.Column;
                Main.instance.heads.heads[nameL0].Resize(resizeValue, false, stopall);
            }
            if (Main.instance.heads.heads[nameL0].childs[nameL1].LastCell.Column < Main.instance.references.references[name].PS.LastColumn.Column)
            {
                int resizeValue = Main.instance.references.references[name].PS.LastColumn.Column - Main.instance.heads.heads[nameL0].childs[nameL1].LastCell.Column;
                Main.instance.heads.heads[nameL0].childs[nameL1].Resize(resizeValue, false, stopall);
            }
            if (Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].LastCell.Column < Main.instance.references.references[name].PS.LastColumn.Column)
            {
                int resizeValue = Main.instance.references.references[name].PS.LastColumn.Column - Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].LastCell.Column;
                Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].Resize(resizeValue, false, stopall);
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
            GlobalMethods.ToLog(this, sender);
            Close();
        }
    }
}
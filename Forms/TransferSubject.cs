//using Microsoft.Office.Interop.Excel;
using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Meter.Forms
{
    partial class TransferSubject : Form
    {
        string? nameL0, nameL1, nameL2;
        string name;
        public string adr;
        ReferenceObject referenceObject;
        public TransferSubject()
        {
            InitializeComponent();
            if (Main.instance.heads.heads != null)
            {
                ComboBox11.DataSource = Main.instance.heads.heads.Keys.ToList();
            }
            referenceObject = RangeReferences.activeTable;
            name = RangeReferences.activeTable._name;
            this.FlowLayoutPanel3.Visible = false;

            if (referenceObject.HeadL2.FirstCell.Column >= ((Excel.Range)referenceObject.PS.Head.Cells[1, 1]).Column)
            {
                this.RBleft.Enabled = false;
            }
            if (referenceObject.HeadL2.LastCell.Column <= ((Excel.Range)referenceObject.PS.Head.Cells[1, referenceObject.PS.Head.Cells.Count]).Column)
            {
                this.RBright.Enabled = false;
            }
            this.RBleft.Checked =  false;
            this.RBright.Checked = false;
            this.RBzone.Checked = false;
        }

        private void TransferSubject_Shown(object sender, EventArgs e)
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

        public void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            // MessageBox.Show(referenceObject.HeadL2.FirstCell.Column + " " + ((Excel.Range)referenceObject.PS.Head.Cells[1, 1]).Column);
        }
        public void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            // MessageBox.Show(referenceObject.HeadL2.LastCell.Column + " " + ((Excel.Range)referenceObject.PS.Head.Cells[1, referenceObject.PS.Head.Cells.Count]).Column);
        }
        public void RadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (this.RBzone.Checked)
            {
                this.FlowLayoutPanel3.Visible = true;
                this.ComboBox11.Text = "";
            }
            else
            {
                this.FlowLayoutPanel3.Visible = false;
                this.ComboBox11.Text = "";
            }
        }

        public void BtnOk_Click(object sender, EventArgs e)
        {
            string nameL0old = referenceObject.HeadL0._name;
            string nameL1old = referenceObject.HeadL1._name;
            string nameL2old = referenceObject.HeadL2._name;

            if (!RBleft.Checked && !RBright.Checked && !RBzone.Checked)
            {
                MessageBox.Show("Выберите один вариант перемещения");
                return;
            }

            if (RBleft.Checked)
            {
                Main.instance.StopAll();
                referenceObject.PS.Head.Resize[1, 1].Offset[0, -1].MergeArea.Resize[1, 1].Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, referenceObject.PS.Range.Cut());
                //Main.instance.wsCh.Range[adr].Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, referenceObject.PS.Range.Cut());
                Main.instance.ResumeAll();
            } 
            else if (RBright.Checked)
            {
                Main.instance.StopAll();
                referenceObject.PS.Head.Offset[0, 1].Resize[1, 1].MergeArea.Offset[0, 1].Resize[1, 1].Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, referenceObject.PS.Range.Cut());
                //Main.instance.wsCh.Range[adr].Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, referenceObject.PS.Range.Cut());
                Main.instance.ResumeAll();
            }
            else if (RBzone.Checked)
            {
                string newName = referenceObject._name.Replace(nameL2old, nameL2);
                string oldName = referenceObject._name;
                if (string.IsNullOrEmpty(nameL0) || string.IsNullOrEmpty(nameL1) || string.IsNullOrEmpty(nameL2))
                {
                    MessageBox.Show("Введите расположение субъекта в таблице!");
                    return;
                }
                if (Main.instance.heads.heads == null || !Main.instance.heads.heads.ContainsKey(nameL0))
                {
                    MessageBox.Show("Введите существующее расположение субъекта в таблице!");
                    return;
                }
                if (Main.instance.heads.heads[nameL0].childs == null || !Main.instance.heads.heads[nameL0].childs.ContainsKey(nameL1))
                {
                    MessageBox.Show("Введите существующее расположение субъекта в таблице!");
                    return;
                }
                if (Main.instance.heads.heads[nameL0].childs[nameL1].childs == null || !Main.instance.heads.heads[nameL0].childs[nameL1].childs.ContainsKey(nameL2))
                {
                    MessageBox.Show("Введите существующее расположение субъекта в таблице!");
                    return;
                }
                if (nameL0old == nameL0 && nameL1old == nameL1 && nameL2old == nameL2)
                {
                    MessageBox.Show("Субъект уже в этой области!");
                    return;
                }

                if (Main.instance.references.references.ContainsKey(newName))
                {
                    MessageBox.Show("Такой субъект в выбранной зоне уже существует!");
                    return;
                }
                if (referenceObject.PS.Head.Columns.Count == referenceObject.HeadL2.Range.Columns.Count)
                {
                    MessageBox.Show("Невозможно переместить единственный субъект в зоне!");
                    return;
                }

                Main.instance.StopAll();
                int cols = referenceObject.PS.Head.Columns.Count;

                Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].Range.Resize[1, 1].Offset[1, 0].MergeArea.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, referenceObject.PS.Range.Cut());
                
                Main.instance.heads.heads[nameL0old].Decrease(cols, false);
                Main.instance.heads.heads[nameL0old].childs[nameL1old].Decrease(cols, false);
                Main.instance.heads.heads[nameL0old].childs[nameL1old].childs[nameL2old].Decrease(cols, false);
                
                Main.instance.heads.heads[nameL0].Increase(cols, false);
                Main.instance.heads.heads[nameL0].childs[nameL1].Increase(cols, false);
                Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].Increase(cols, false);

                Main.instance.formulas.formulas.Values.SelectMany(f => f).Where(t => t.text.Contains(oldName)).ToList().ForEach(t => t.text = t.text.Replace(oldName, newName));

                foreach (ChildObject co in referenceObject.childs.Values)
                {
                    Excel.Range r1 = (Excel.Range)co.Head.Cells[1, 1];
                    co._name = newName;
                    r1.Value = newName;
                }

                Main.instance.references.references.Remove(referenceObject._name);
                Main.instance.references.references.Add(newName, referenceObject);
                Main.instance.references.references[newName]._name = newName;

                Main.instance.ResumeAll();

                // MessageBox.Show(Main.instance.heads.heads[nameL0].rangeAddress + " " + Main.instance.heads.heads[nameL0].childs[nameL1].rangeAddress + " " + Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].rangeAddress);
            }
            referenceObject.PS.Head.Select();
            Close();
            #region Old
                // GlobalMethods.ToLog(this, sender);
                // // name = TextBox15.Text + " " + ComboBox13.Text;
                // if (string.IsNullOrEmpty(nameL0) || string.IsNullOrEmpty(nameL1) || string.IsNullOrEmpty(nameL2))
                // {
                //     MessageBox.Show("Введите расположение субъекта в таблиуе!");
                //     return;
                // }
    
                // if (Main.instance.heads.heads != null && Main.instance.heads.heads.ContainsKey(nameL0))
                // {
                //     if (Main.instance.heads.heads[nameL0].childs != null && Main.instance.heads.heads[nameL0].childs.ContainsKey(nameL1))
                //     {
                //         if (Main.instance.heads.heads[nameL0].childs[nameL1].childs != null && Main.instance.heads.heads[nameL0].childs[nameL1].childs.ContainsKey(nameL2))
                //         {
                //             Main.instance.StopAll();
                //             adr = Main.instance.heads.heads[nameL0].childs[nameL1].childs[nameL2].LastCell.Offset[1, 1].Address[false, false];
                //             Transfer();
                //             UpdateUps(false);
                //             Main.instance.ResumeAll();
                //             Close();
                //         }
                //         else if (Main.instance.heads.heads[nameL0].childs[nameL1].childs == null)
                //         {
                //             Main.instance.heads.heads[nameL0].childs[nameL1].childs = new Dictionary<string, HeadObject> ();
                //             Main.instance.StopAll();
                //             Excel.Range r = Main.instance.heads.heads[nameL0].childs[nameL1].LastCell.Offset[1, 0].Resize[1];
                //             r.Value = nameL2;
                //             HeadObject ho = new HeadObject()
                //             {
                //                 WS = Main.instance.wsCh,
                //                 _name = nameL2,
                //                 Range = r,
                //                 _level = Level.level2,
                //                 parentID = Main.instance.heads.heads[nameL0].childs[nameL1].ID,
                //             };
                //             Main.instance.heads.heads[nameL0].childs[nameL1].childs.Add(nameL2, ho);
                //             if (!HeadReferences.idDictionary.ContainsKey(ho.ID)) HeadReferences.idDictionary.Add(ho.ID, ho);
                //             adr = r.Offset[1].Address[false, false];
                //             Transfer(false);
                //             UpdateUps(false);
                //             Main.instance.ResumeAll();
                //             Close();
                //         }
                //         else
                //         {
                //             Main.instance.StopAll();
                //             Excel.Range r = Main.instance.heads.heads[nameL0].childs[nameL1].LastCell.Offset[-1, 1].Resize[42];
                //             string address = Main.instance.heads.heads[nameL0].childs[nameL1].LastCell.Offset[1, 1].Resize[1].Address[false, false];
                //             r.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                //             r = Main.instance.wsCh.Range[address];
                //             r.Value = nameL2;
                //             HeadObject ho = new HeadObject()
                //             {
                //                 WS = Main.instance.wsCh,
                //                 _name = nameL2,
                //                 Range = r,
                //                 _level = Level.level2,
                //                 parentID = Main.instance.heads.heads[nameL0].childs[nameL1].ID,
                //             };
                //             Main.instance.heads.heads[nameL0].childs[nameL1].childs.Add(nameL2, ho);
                //             if (!HeadReferences.idDictionary.ContainsKey(ho.ID)) HeadReferences.idDictionary.Add(ho.ID, ho);
                //             adr = r.Offset[1].Address[false, false];
                //             Transfer(false);
                //             UpdateUps(false);
                //             Main.instance.ResumeAll();
                //             Close();
                //         }
                //     }
                //     else if (Main.instance.heads.heads[nameL0].childs == null)
                //     {
                //         Main.instance.heads.heads[nameL0].childs = new Dictionary<string, HeadObject>();
                //         Main.instance.StopAll();
                //         Excel.Range r = Main.instance.heads.heads[nameL0].LastCell.Offset[1, 0].Resize[1];
                //         r.Value = nameL1;
                //         HeadObject ho = new HeadObject()
                //         {
                //             WS = Main.instance.wsCh,
                //             _name = nameL1,
                //             Range = r,
                //             _level = Level.level1,
                //             childs = new Dictionary<string, HeadObject>(),
                //             parentID = Main.instance.heads.heads[nameL0].ID,
                //         };
                //         Main.instance.heads.heads[nameL0].childs.Add(nameL1, ho);
                //         if (!HeadReferences.idDictionary.ContainsKey(ho.ID)) HeadReferences.idDictionary.Add(ho.ID, ho);
                //         r = r.Offset[1];
                //         r.Value = nameL2;
                //         ho = new HeadObject()
                //         {
                //             WS = Main.instance.wsCh,
                //             _name = nameL2,
                //             Range = r,
                //             _level = Level.level2,
                //             parentID = Main.instance.heads.heads[nameL0].childs[nameL1].ID,
                //         };
                //         Main.instance.heads.heads[nameL0].childs[nameL1].childs.Add(nameL2, ho);
                //         if (!HeadReferences.idDictionary.ContainsKey(ho.ID)) HeadReferences.idDictionary.Add(ho.ID, ho);
                //         adr = r.Offset[1].Address[false, false];
                //         Transfer(false);
                //         UpdateUps(false);
                //         Main.instance.ResumeAll();
                //         Close();
                //     }
                //     else
                //     {
                //         Main.instance.StopAll();
                //         Excel.Range r = Main.instance.heads.heads[nameL0].LastCell.Offset[-1, 1].Resize[42];
                //         string address = Main.instance.heads.heads[nameL0].LastCell.Offset[1, 1].Resize[1].Address[false, false];
                //         r.Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                //         r = Main.instance.wsCh.Range[address];
                //         r.Value = nameL1;
                //         HeadObject ho = new HeadObject()
                //         {
                //             WS = Main.instance.wsCh,
                //             _name = nameL1,
                //             Range = r,
                //             _level = Level.level1,
                //             childs = new Dictionary<string, HeadObject>(),
                //             parentID = Main.instance.heads.heads[nameL0].ID,
                //         };
                //         Main.instance.heads.heads[nameL0].childs.Add(nameL1, ho);
                //         if (!HeadReferences.idDictionary.ContainsKey(ho.ID)) HeadReferences.idDictionary.Add(ho.ID, ho);
                //         r = r.Offset[1];
                //         r.Value = nameL2;
                //         ho = new HeadObject()
                //         {
                //             WS = Main.instance.wsCh,
                //             _name = nameL2,
                //             Range = r,
                //             _level = Level.level2,
                //             parentID = Main.instance.heads.heads[nameL0].childs[nameL1].ID,
                //         };
                //         Main.instance.heads.heads[nameL0].childs[nameL1].childs.Add(nameL2, ho);
                //         if (!HeadReferences.idDictionary.ContainsKey(ho.ID)) HeadReferences.idDictionary.Add(ho.ID, ho);
                //         adr = r.Offset[1].Address[false, false];
                //         Transfer(false);
                //         UpdateUps(false);
                //         Main.instance.ResumeAll();
                //         Close();
                //     }
                // }
                // else if (Main.instance.heads.heads == null)
                // {
                //     Main.instance.heads.heads = new Dictionary<string, HeadObject>();
                //     int row = 8;
                //     Main.instance.StopAll();
                //     Excel.Range r = (Excel.Range)Main.instance.wsCh.Cells[row, 2];
                //     r.Value = nameL0;
                //     HeadObject ho = new HeadObject()
                //     {
                //         WS = Main.instance.wsCh,
                //         _name = nameL0,
                //         Range = r,
                //         _level = Level.level0,
                //         childs = new Dictionary<string, HeadObject>(),
                //         parentID = null,
                //     };
                //     Main.instance.heads.heads.Add(nameL0, ho);
                //     if (!HeadReferences.idDictionary.ContainsKey(ho.ID)) HeadReferences.idDictionary.Add(ho.ID, ho);
                //     r = r.Offset[1];
                //     r.Value = nameL1;
                //     ho = new HeadObject()
                //     {
                //         WS = Main.instance.wsCh,
                //         _name = nameL1,
                //         Range = r,
                //         _level = Level.level1,
                //         childs = new Dictionary<string, HeadObject>(),
                //         parentID = Main.instance.heads.heads[nameL0].ID,
                //     };
                //     Main.instance.heads.heads[nameL0].childs.Add(nameL1, ho);
                //     if (!HeadReferences.idDictionary.ContainsKey(ho.ID)) HeadReferences.idDictionary.Add(ho.ID, ho);
                //     r = r.Offset[1];
                //     r.Value = nameL2;
                //     ho = new HeadObject()
                //     {
                //         WS = Main.instance.wsCh,
                //         _name = nameL2,
                //         Range = r,
                //         _level = Level.level2,
                //         parentID = Main.instance.heads.heads[nameL0].childs[nameL1].ID,
                //     };
                //     Main.instance.heads.heads[nameL0].childs[nameL1].childs.Add(nameL2, ho);
                //     if (!HeadReferences.idDictionary.ContainsKey(ho.ID)) HeadReferences.idDictionary.Add(ho.ID, ho);
                //     adr = r.Offset[1].Address[false, false];
                //     Transfer(false);
                //     UpdateUps(false);
                //     Main.instance.ResumeAll();
                //     Close();
                // }
                // else
                // {
                //     int row = Main.instance.heads.heads.Values.Select(n => n.Range.Row).Max();
                //     row = row + 43;
                //     Main.instance.StopAll();
                //     Excel.Range r = (Excel.Range)Main.instance.wsCh.Cells[row, 2];
                //     r.Value = nameL0;
                //     HeadObject ho = new HeadObject()
                //     {
                //         WS = Main.instance.wsCh,
                //         _name = nameL0,
                //         Range = r,
                //         _level = Level.level0,
                //         childs = new Dictionary<string, HeadObject>(),
                //         parentID = null,
                //     };
                //     Main.instance.heads.heads.Add(nameL0, ho);
                //     if (!HeadReferences.idDictionary.ContainsKey(ho.ID)) HeadReferences.idDictionary.Add(ho.ID, ho);
                //     r = r.Offset[1];
                //     r.Value = nameL1;
                //     ho = new HeadObject()
                //     {
                //         WS = Main.instance.wsCh,
                //         _name = nameL1,
                //         Range = r,
                //         _level = Level.level1,
                //         childs = new Dictionary<string, HeadObject>(),
                //         parentID = Main.instance.heads.heads[nameL0].ID,
                //     };
                //     Main.instance.heads.heads[nameL0].childs.Add(nameL1, ho);
                //     if (!HeadReferences.idDictionary.ContainsKey(ho.ID)) HeadReferences.idDictionary.Add(ho.ID, ho);
                //     r = r.Offset[1];
                //     r.Value = nameL2;
                //     ho = new HeadObject()
                //     {
                //         WS = Main.instance.wsCh,
                //         _name = nameL2,
                //         Range = r,
                //         _level = Level.level2,
                //         parentID = Main.instance.heads.heads[nameL0].childs[nameL1].ID,
                //     };
                //     Main.instance.heads.heads[nameL0].childs[nameL1].childs.Add(nameL2, ho);
                //     if (!HeadReferences.idDictionary.ContainsKey(ho.ID)) HeadReferences.idDictionary.Add(ho.ID, ho);
                //     adr = r.Offset[1].Address[false, false];
                //     Transfer(false);
                //     UpdateUps(false);
                //     Main.instance.ResumeAll();
                //     Close();
                // }
            #endregion
        }

        private void Transfer(bool insert = true)
        {
            // referenceObject.PS.Range.Select();
            // Main.instance.wsCh.Range[adr].Insert(Shift: Excel.XlInsertShiftDirection.xlShiftToRight, referenceObject.PS.Range.Cut());
            // Main.instance.references.CreateNew(name, types[0], adr, insert, false);
            // types.RemoveAt(0);
            // if (types.Count > 0)
            // {
            //     foreach (string t in types)
            //     {
            //         Main.instance.references.references[name].AddNewDBL1StandartOther(t, false);
            //         Main.instance.references.references[name].AddNewPS(t, "ручное", false);
            //         Main.instance.references.references[name].PS.childs[t].childs["ручное"].ChangeCod();
            //     }
            // }
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
using FluentDragDrop;
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
    public partial class FormulaEditor : Form
    {
        ReferenceObject referenceObject;
        string nameL1;
        string myID;

        private static Random _random = new Random();
        List<ListViewItem> list = new List<ListViewItem>();
        List<Control> open = new List<Control>();
        HashSet<Control> active = new HashSet<Control>();

        ForTags lastControl = null;
        string oldFormula = "";
        string newFormula = "";
        //ButtonsType? lastType = null;

        public FormulaEditor(ref ReferenceObject referenceObject, string nameL1)
        {
            this.referenceObject = referenceObject;
            this.nameL1 = nameL1;
            myID = referenceObject.DB.childs[nameL1].ID;
            InitializeComponent();
        }

        private void FormulaEditor_Load(object sender, EventArgs e)
        {

        }

        private void FormulaEditor_Shown(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this);
            ListViewItem item;
            listView1.Columns[0].Width = listView1.Width;
            this.Text = referenceObject._name;

            foreach (ReferenceObject ro in Main.instance.references.references.Values)
            {
                foreach (ChildObject co in ro.DB.childs.Values)
                {
                    if (co.ID == referenceObject.DB.childs[nameL1].ID)
                    {
                        continue;
                    }
                    // if (co._name == "план")
                    // {
                    //     continue;
                    // }
                    item = new ListViewItem();
                    item.Text = ro._name + " " + co._name;
                    item.Tag = new ForTags();
                    if (co._name == "план")
                    {
                        item.BackColor = ColorTranslator.FromHtml("#CCFF33");
                    }
                    ((ForTags)item.Tag).ID = co.childs["основное"].ID;
                    ((ForTags)item.Tag).type = ButtonsType.subject;
                    listView1.Items.Add(item);
                    list.Add(item);
                }
            }

            if (Main.instance.formulas.formulas.ContainsKey(myID))
            {
                foreach (ForTags f in Main.instance.formulas.formulas[myID])
                {
                    Button b = new Button();
                    b.Tag = f;
                    if (f.type == ButtonsType.subject)
                    {
                        b.Text = f.text;
                        b.Font = new Font(b.Font, FontStyle.Bold);
                        b.UseVisualStyleBackColor = true;
                        b.AutoSize = true;
                        b.MouseDown += control_MouseDown;   
                        SubjectContextMenu(b);
                    }
                    else if (f.type == ButtonsType.constant)
                    {
                        b.Text = f.text;
                        b.Font = new Font(b.Font, FontStyle.Bold);
                        b.UseVisualStyleBackColor = true;
                        b.AutoSize = true;
                        b.MouseDown += control_MouseDown;
                        b.MouseDown += EnterConstant;
                    }
                    else
                    {
                        b.Text = f.text;
                        b.Font = new Font(b.Font, FontStyle.Bold);
                        b.UseVisualStyleBackColor = true;
                        b.Width = button5.Width;
                        b.Height = button5.Height;
                        b.MouseDown += control_MouseDown;
                        if (b.Text == ")" || b.Text == "(")
                        {
                            b.MouseDown += ShowPara;
                        }
                    }
                    oldFormula += "{" + b.Text + "} ";
                    flowLayoutPanel1.Controls.Add(b);
                }
                UpdateCheck();
            }

            if (!NewMenuBase.editedFormulas.Contains(myID))
            {
                NewMenuBase.editedFormulas.Add(myID);
            }
            else
            {
                Close();
            }
        }

        private void listView1_MouseDown(object sender, MouseEventArgs e)
        {
            listView1.InitializeDragAndDrop()
            .Copy()
            .OnMouseMove()
            .If(() => listView1.SelectedItems.OfType<ListViewItem>().ToArray().Any())
            .WithData(() => listView1.SelectedItems.OfType<ListViewItem>().ToArray())
            .WithPreview((bitmap) =>
            {
                Bitmap bmp = new Bitmap(75, 23);
                Button b = new Button();
                b.Text = listView1.SelectedItems[0].Text;
                b.DrawToBitmap(bmp, new Rectangle(0, 0, 75, 23));
                return bmp;
            }).BehindCursor().To(flowLayoutPanel1, (flowLayoutPanel1, btn) =>
            {
                Button b = new Button();
                b.Text = listView1.SelectedItems[0].Text;
                b.Font = new Font(b.Font, FontStyle.Bold);
                b.Tag = listView1.SelectedItems[0].Tag;
                b.SpecialTag().text = b.Text;
                b.UseVisualStyleBackColor = true;
                b.AutoSize = true;
                b.MouseDown += control_MouseDown;

                Point newPosition = flowLayoutPanel1.PointToClient(MousePosition);

                int index = flowLayoutPanel1.Controls.Count;
                int newIndex = GetNewIndex(b, newPosition);

                flowLayoutPanel1.Controls.Add(b);
                if (newIndex != index)
                {
                    flowLayoutPanel1.Controls.SetChildIndex(b, newIndex);
                }

                if (b.SpecialTag().type == ButtonsType.subject)
                {
                    SubjectContextMenu(b);
                }

                UpdateCheck();
            });
        }

        private void listView1_Click(object sender, EventArgs e)
        {
            foreach (Control item in flowLayoutPanel1.Controls)
            {
                if (item is Button button)
                {
                    try
                    {
                        if (listView1.SelectedItems[0].Text == button.Text)
                        {
                            button.BackColor = SystemColors.Highlight;
                        }
                        else
                        {
                            button.BackColor = SystemColors.Control;
                        }
                    }
                    catch (System.Exception)
                    {
                        button.BackColor = SystemColors.Control;
                    }
                }
            }
        }

        private void control_MouseDown(object sender, MouseEventArgs e)
        {
            var btn = (Button)sender;

            btn.InitializeDragAndDrop()
                .Move()
                .OnMouseMove()
                .WithData(() => btn)
                .OnBeforeStart(() =>
                {
                    trash.Visible = true;
                })
                .WithPreview().RelativeToCursor().OnCancel(() =>
                {
                    trash.Visible = false;
                }).To(flowLayoutPanel1, (flowLayoutPanel1, btn) =>
                {
                    Point newPosition = flowLayoutPanel1.PointToClient(MousePosition);

                    int index = flowLayoutPanel1.Controls.GetChildIndex(btn);
                    int newIndex = GetNewIndex(btn, newPosition);

                    if (newIndex != index)
                    {
                        flowLayoutPanel1.Controls.SetChildIndex(btn, newIndex);
                    }

                    trash.Visible = false;

                    UpdateCheck();
                });
        }

        private void control1_MouseDown(object sender, MouseEventArgs e)
        {
            var btn = (Button)sender;

            btn.InitializeDragAndDrop()
                .Copy()
                .Immediately()
                .WithData(btn)
                .WithPreview().RelativeToCursor().To(flowLayoutPanel1, (flowLayoutPanel1, btn) =>
                {
                    Button b = new Button();
                    b.Text = btn.Text;
                    b.Font = new Font(b.Font, FontStyle.Bold);
                    b.Tag = new ForTags();
                    b.UseVisualStyleBackColor = true;
                    b.MouseDown += control_MouseDown;

                    Point newPosition = flowLayoutPanel1.PointToClient(MousePosition);

                    int index = flowLayoutPanel1.Controls.Count;
                    int newIndex = GetNewIndex(b, newPosition);

                    flowLayoutPanel1.Controls.Add(b);
                    if (newIndex != index)
                    {
                        flowLayoutPanel1.Controls.SetChildIndex(b, newIndex);
                    }

                    if (b.Text == "(" || b.Text == ")")
                    {
                        b.Width = btn.Width;
                        b.Height = btn.Height;
                        b.AutoSize = true;
                        b.SpecialTag().text = b.Text;
                        b.SpecialTag().type = ButtonsType.scobki;
                        b.MouseDown += ShowPara;
                    }
                    else if (b.Text == "+" || b.Text == "-" || b.Text == "*" || b.Text == "/")
                    {
                        b.Width = btn.Width;
                        b.Height = btn.Height;
                        b.AutoSize = true;
                        b.SpecialTag().text = b.Text;
                        b.SpecialTag().type = ButtonsType.znaki;
                    }
                    else
                    {
                        b.Text = "0";
                        b.SpecialTag().text = b.Text;
                        b.AutoSize = true;
                        b.SpecialTag().type = ButtonsType.constant;
                        b.MouseDown += EnterConstant;
                    }

                    UpdateCheck();
                    
                });
        }

        private void OpenFormula(Control button)
        {
            //if (e.Button == MouseButtons.Right)
            {
                string newID = button.SpecialTag().ID;
                ChildObject co = (ChildObject)RangeReferences.idDictionary[newID];
                string newNameL1 = co.GetParent<ChildObject>()._name;
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
        }

        private int GetNewIndex(Control control, Point newPosition)
        {
            int index;
            if (control == null)
            {
                index = flowLayoutPanel1.Controls.Count;
            }
            else
            {
                if (flowLayoutPanel1.Controls.Contains(control))
                {
                    index = flowLayoutPanel1.Controls.GetChildIndex(control);
                }
                else
                {
                    index = flowLayoutPanel1.Controls.Count;
                }
            }
            int controlsCount = flowLayoutPanel1.Controls.Count;

            for (int i = 0; i < controlsCount; i++)
            {
                if (i == index)
                {
                    continue;
                }

                Control otherControl = flowLayoutPanel1.Controls[i];
                int otherTop = otherControl.Top;
                int otherDown = otherControl.Bottom;
                int otherCenter = otherControl.Left + (otherControl.Width / 2);

                if (otherTop <= newPosition.Y && otherDown >= newPosition.Y && newPosition.X < otherCenter)
                {
                    if (i > index)
                    {
                        return i - 1;
                    }
                    else
                    {
                        return i;
                    }
                }
                else if (i == controlsCount - 1 && otherTop <= newPosition.Y && otherDown >= newPosition.Y && newPosition.X > otherCenter)
                {
                    if (flowLayoutPanel1.Controls.Contains(control))
                    {
                        return i;
                    }
                    else
                    {
                        return i + 1;
                    }
                }
            }
            return index;
        }

        private void RegexSearch()
        {
            RegexOptions ro = checkBox1.Checked ? RegexOptions.None : RegexOptions.IgnoreCase;
            if (this.tbSearch.Text != "")
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    try
                    {
                        listView1.Items.Clear();
                        string search = this.tbSearch.Text;
                        search = search.Replace("*", @".*");

                        var deviceIds = list.AsEnumerable();
                        var matchingIds = deviceIds.Where(id => Regex.IsMatch(id.Text, pattern: search, ro)).ToArray();
                        listView1.Items.AddRange(matchingIds.ToArray());
                    }
                    catch (ArgumentException)
                    {
                        this.Invoke((MethodInvoker)(() =>
                        {
                            listView1.Items.Clear();
                            listView1.Items.AddRange(list.OrderBy(m => m.Text).ToArray());
                        }));
                    }

                }));
            }
            else
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    listView1.Items.Clear();
                    listView1.Items.AddRange(list.OrderBy(m => m.Text).ToArray());
                }));
            }

        }

        private void RegexSearchFormula()
        {
            RegexOptions ro = checkBox1.Checked ? RegexOptions.None : RegexOptions.IgnoreCase;
            if (this.tbSearch.Text != "" && cbSearchFormula.Checked == true)
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    try
                    {
                        string search = this.tbSearch.Text;
                        search = search.Replace("*", @".*");
                        foreach (Control item in flowLayoutPanel1.Controls)
                        {
                            if (item is Button button)
                            {
                                bool match = Regex.IsMatch(button.Text, pattern: search, ro);
                                button.BackColor = match ? Color.LightSkyBlue : SystemColors.Control;
                            }
                        }
                    }
                    catch (ArgumentException)
                    {
                        this.Invoke((MethodInvoker)(() =>
                        {
                            foreach (Control item in flowLayoutPanel1.Controls)
                            {
                                item.BackColor = SystemColors.Control;
                            }
                        }));
                    }
                }));
            }
            else
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    foreach (Control item in flowLayoutPanel1.Controls)
                    {
                        item.BackColor = SystemColors.Control;
                    }
                }));
            }
        }

        private void tbSearch_TextChanged(object sender, EventArgs e)
        {
            RegexSearch();
            RegexSearchFormula();
        }

        private void cbSearchFormula_CheckedChanged(object sender, EventArgs e)
        {
            RegexSearchFormula();
        }

        private void UpdateCheck()
        {
            int myIndex;
            lastControl = null;
            //lastType = null;
            open.Clear();
            foreach (Control item in flowLayoutPanel1.Controls)
            {
                myIndex = flowLayoutPanel1.Controls.GetChildIndex(item);
                switch (item.SpecialTag().type)
                {
                    case ButtonsType.subject:
                        if ((item.SpecialTag().text == "#ссылка" || item.SpecialTag().ID == null || item.Text.Contains("удален")) || (lastControl != null && lastControl.type != ButtonsType.znaki))
                        {
                            item.BackColor = Color.Red;
                        }
                        else
                        {
                            item.BackColor = SystemColors.ButtonFace;
                        }

                        lastControl = item.SpecialTag();
                        //lastType = ButtonsType.subject;
                        break;
                    case ButtonsType.scobki:

                        UpdateScobki(item);

                        break;
                    case ButtonsType.znaki:
                        if ((lastControl != null && lastControl.type != ButtonsType.constant && lastControl.type != ButtonsType.subject) || myIndex == flowLayoutPanel1.Controls.Count - 1)
                        {
                            item.BackColor = Color.Red;
                        }
                        else
                        {
                            item.BackColor = SystemColors.ButtonFace;
                        }

                        lastControl = item.SpecialTag();
                        //lastType= ButtonsType.znaki;
                        break;
                    case ButtonsType.constant:
                        if (lastControl != null && lastControl.type != ButtonsType.znaki)
                        {
                            item.BackColor = Color.Red;
                        }
                        else
                        {
                            item.BackColor = SystemColors.ButtonFace;
                        }

                        lastControl = item.SpecialTag();
                        //lastType = ButtonsType.constant;
                        break;
                    default:
                        break;
                }
            }

            if (open.Count > 0)
            {
                foreach (Control item in open)
                {
                    ((ForTags)item.Tag).ID = "";
                    item.ForeColor = SystemColors.ControlText;
                    item.BackColor = Color.Red;
                }
            }
        }

        private void UpdateScobki(Control item)
        {
            int myIndex = flowLayoutPanel1.Controls.GetChildIndex(item);
            if (item.Text == "(")
            {
                open.Add(item);
                ((ForTags)item.Tag).ID = Guid.NewGuid().ToString();
            }
            else if (item.Text == ")")
            {
                if (open.Count > 0)
                {
                    Color randomColor;

                    randomColor = GetRandomColor();

                    open[open.Count - 1].ForeColor = randomColor;
                    open[open.Count - 1].BackColor = SystemColors.ButtonFace;

                    item.ForeColor = randomColor;
                    item.BackColor = SystemColors.ButtonFace;

                    ((ForTags)item.Tag).ID = ((ForTags)open[open.Count - 1].Tag).ID;

                    open.RemoveAt(open.Count - 1);

                    if (myIndex != 0)
                    {
                        if (flowLayoutPanel1.Controls[myIndex - 1].SpecialTag().type == ButtonsType.znaki)
                        {
                            item.BackColor = Color.Red;
                        }
                    }
                }
                else
                {
                    ((ForTags)item.Tag).ID = "";
                    item.ForeColor = SystemColors.ControlText;
                    item.BackColor = Color.Red;
                }
            }
        }

        private async void ShowPara(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                Control control = (Control)sender;
                string id = ((ForTags)control.Tag).ID;

                int interval = 400; // миллисекунды
                int blinkCount = 4; // количество мерцаний

                foreach (Control item in flowLayoutPanel1.Controls)
                {
                    if (((ForTags)item.Tag).ID == id && control != item && !active.Contains(item))
                    {
                        Color originalColor = SystemColors.ButtonFace;
                        Color blinkColor = item.ForeColor;
                        active.Add(item);

                        for (int i = 0; i < blinkCount; i++)
                        {
                            await Task.Delay(interval);
                            item.BackColor = (item.BackColor == originalColor) ? blinkColor : originalColor;
                        }

                        item.BackColor = originalColor;
                        active.Remove(item);
                        return;
                    }
                }
            }
        }

        static Color GetRandomColor(float saturation = 1f, float value = 1f)
        {
            float hue = (float)_random.NextDouble(); // случайный оттенок
            return FromHsv(new HsvColor(hue, saturation, value));
        }

        static Color FromHsv(HsvColor hsv)
        {
            int hi = Convert.ToInt32(Math.Floor(hsv.Hue * 6)) % 6;
            double f = hsv.Hue * 6 - Math.Floor(hsv.Hue * 6);
            double p = hsv.Value * (1 - hsv.Saturation);
            double q = hsv.Value * (1 - f * hsv.Saturation);
            double t = hsv.Value * (1 - (1 - f) * hsv.Saturation);

            byte v = Convert.ToByte(hsv.Value * 255);
            byte p1 = Convert.ToByte(p * 255);
            byte q1 = Convert.ToByte(q * 255);
            byte t1 = Convert.ToByte(t * 255);

            if (hi == 0) return Color.FromArgb(255, v, t1, p1);
            else if (hi == 1) return Color.FromArgb(255, q1, v, p1);
            else if (hi == 2) return Color.FromArgb(255, p1, v, t1);
            else if (hi == 3) return Color.FromArgb(255, p1, q1, v);
            else if (hi == 4) return Color.FromArgb(255, t1, p1, v);
            else return Color.FromArgb(255, v, p1, q1);
        }

        struct HsvColor
        {
            public float Hue;
            public float Saturation;
            public float Value;

            public HsvColor(float hue, float saturation, float value)
            {
                Hue = hue;
                Saturation = saturation;
                Value = value;
            }
        }

        private void trash_DragDrop(object sender, DragEventArgs e)
        {
            Button c = e.Data.GetData(typeof(Button)) as Button;
            if (c != null)
            {
                if (flowLayoutPanel1.Controls.Contains(c))
                {
                    flowLayoutPanel1.Controls.Remove(c);
                }
            }
            UpdateCheck();
        }

        private void trash_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            if (MessageBox.Show("Все внесенные изменения будут отменены. Вы уверены?", "",MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Close();
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            List<ForTags> formulaIDs = new List<ForTags>();
            //List<string> formula = new List<string>();
            string formula = "";
            List<ForTags> formulaForArch = new List<ForTags>();

            foreach (Control item in flowLayoutPanel1.Controls)
            {
                if (item.BackColor == Color.Red)
                {
                    MessageBox.Show("В формуле есть ошибки. Проверьте правильность введенной формулы!");
                    return;
                }
                formulaForArch.Add(item.SpecialTag());
                if (item.SpecialTag().type == ButtonsType.subject)
                {
                    string adr;
                    Excel.Range r = null;
                    if (item.SpecialTag().ID != null)
                    {
                        r = ((Excel.Range)((ChildObject)RangeReferences.idDictionary[item.SpecialTag().ID]).Body.Cells[1, 1]);
                        adr = r.Address[false, false];
                    }
                    else
                    {
                        adr = "#ссылка";
                    }
                    
                    formula += adr;
                    if (r != null) Marshal.ReleaseComObject(r);
                }
                else
                {
                    formula += item.Text;
                }
                newFormula += "{" + item.Text + "} ";
            }
            formula = formula.Replace(",",".");
            // if (formula == "=")
            // {
            //     formula = "";
            // }
            if (Main.instance.formulas.formulas.ContainsKey(myID))
            {
                Main.instance.formulas.formulas[myID].Clear();
                Main.instance.formulas.formulas[myID].AddRange(formulaForArch.ToArray());
            }
            else
            {
                Main.instance.formulas.formulas.Add(myID, new List<ForTags>());
                Main.instance.formulas.formulas[myID].AddRange(formulaForArch.ToArray());
            }

            referenceObject.DB.childs[nameL1].WriteFormula(formula);

            GlobalMethods.ToLog("Изменена формула для субъекта {" + referenceObject._name + "} " + nameL1 + " с '" + oldFormula + "' на '" + newFormula + "'");

            //MessageBox.Show(formula + "; Корректность:" + Main.instance.xlApp.WorksheetFunction.IsError(formula) + "; результат:" + Main.instance.xlApp.Evaluate(formula));
            Close();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();
        }
        private void EnterConstant(object sender, MouseEventArgs e)
        {
            Control b = (Control)sender;
            if (e.Button == MouseButtons.Right)
            {
                using (EnterCoef form = new EnterCoef(b))
                {
                    form.ShowDialog();
                    b.Text = form.newVal;
                    b.SpecialTag().text = b.Text;
                }
                
            }
            //flowLayoutPanel1.Controls.Clear();
        }
        private void FormulaEditor_Closing(object sender, EventArgs e)
        {
            NewMenuBase.editedFormulas.Remove(myID);
        }

        private void SubjectContextMenu(Control b)
        {
            ToolStripItem tsi = null;

            ContextMenuStrip menu = new ContextMenuStrip();

            if (b.SpecialTag().ID != null)
            {
                tsi = menu.Items.Add("Показать в счетчиках");

                tsi.Click += (object sender, EventArgs e) =>
                {
                    string newID = b.SpecialTag().ID;
                    ChildObject co = (ChildObject)RangeReferences.idDictionary[newID];
                    ReferenceObject ro = co.GetFirstParent;
                    ro.PS.Range.Select();
                };

                string id = b.SpecialTag().ID;
                if (RangeReferences.idDictionary.ContainsKey(id))
                {
                    if (((ChildObject)RangeReferences.idDictionary[id]).GetParent<ChildObject>().HasItem("формула"))
                    {
                        tsi = menu.Items.Add("Показать формулу");
                        tsi.Click += (object sender, EventArgs e) =>
                        {
                            OpenFormula(b);
                        };
                    }
                }
            }
            b.ContextMenuStrip = menu;
        }
    
        private void Test(object sender, EventArgs e) 
        {
            Control b = (Control)sender;
            MessageBox.Show(b.SpecialTag().ID);
        }
    }
}

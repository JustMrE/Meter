using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Meter
{

    partial class Rename : Form
    {
        ReferenceObject referenceObject;
        HeadObject headObject;
        string oldnameShown, oldName, newName;
        public Rename(ReferenceObject referenceObject)
        {
            this.headObject = null;
            this.referenceObject = referenceObject;
            oldnameShown = referenceObject._name.Replace(" " + referenceObject.HeadL2._name,"");
            oldName = referenceObject._name;
            InitializeComponent();
        }
        public Rename(HeadObject headObject)
        {
            this.referenceObject = null;
            this.headObject = headObject;
            oldnameShown = headObject._name;
            oldName = headObject._name;
            InitializeComponent();
        }

        private void Rename_Shown(object sender, System.EventArgs e)
        {
            GlobalMethods.ToLog(this);
            tbNewName.Text = oldnameShown;
        }

        protected void btnOk_Click(object sender, System.EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            if (referenceObject != null) RenameSubject();
            else if (headObject != null) RenameHead();
            GlobalMethods.ToLog("Изменено название субъекта с '" + oldName + "' на '" + newName + "'");
            Close();
        }

        private void RenameSubject()
        {
            newName = tbNewName.Text + " " + referenceObject.HeadL2._name;

            if (Main.instance.references.references.ContainsKey(newName))
            {
                MessageBox.Show("Это имя уже используется! Выберите другое.");
                return;
            }

            Main.instance.references.references.Remove(oldName);
            Main.instance.references.references.Add(newName, referenceObject);

            Main.instance.StopAll();
            referenceObject._name = newName;
            foreach (ChildObject co in referenceObject.childs.Values)
            {
                Excel.Range r = ((Excel.Range)co.Head.Cells[1, 1]);
                co._name = newName;
                r.Value = newName;
                Marshal.ReleaseComObject(r);
            }
            Main.instance.ResumeAll();
        }

        private void RenameHead()
        {
            newName = tbNewName.Text;
            if (Main.instance.heads.HasHead(newName))
            {
                MessageBox.Show("Это имя уже используется! Выберите другое.");
                return;
            }
            if (headObject._level != Level.level0)
            {
                headObject.GetParent.childs.Remove(oldName);
                headObject.GetParent.childs.Add(newName, headObject);
            }
            else
            {
                Main.instance.heads.heads.Remove(oldName);
                Main.instance.heads.heads.Add(newName, headObject);
            }

            Main.instance.StopAll();
            headObject._name = newName;
            Excel.Range r = ((Excel.Range)headObject.Range.Cells[1, 1]);
            r.Value = newName;
            Marshal.ReleaseComObject(r);
            if (headObject._level == Level.level2)
            {
                List<string> subjects = Main.instance.references.references.Keys.Where(k => k.Contains(oldName)).ToList();
                foreach (string n in subjects)
                {
                    ReferenceObject ro = Main.instance.references.references[n];
                    string name = ro._name.Replace(oldName, newName);

                    foreach (ChildObject co in ro.childs.Values)
                    {
                        Excel.Range r1 = ((Excel.Range)co.Head.Cells[1, 1]);
                        co._name = name;
                        r1.Value = name;
                        Marshal.ReleaseComObject(r);
                    }

                    Main.instance.references.references.Remove(ro._name);
                    Main.instance.references.references.Add(name, ro);
                }
            }
            
            Main.instance.ResumeAll();
        }

        protected void btnCancel_Click(object sender, System.EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            Close();
        }
    }
}
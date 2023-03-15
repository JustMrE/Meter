using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Meter
{

    partial class Rename : Form
    {
        ReferenceObject referenceObject;
        string oldName, newName;
        public Rename(ReferenceObject referenceObject)
        {
            this.referenceObject = referenceObject;
            oldName = referenceObject._name;
            InitializeComponent();
        }

        private void Rename_Shown(object sender, System.EventArgs e)
        {
            GlobalMethods.ToLog(this);
            tbNewName.Text = oldName;
        }

        protected void btnOk_Click(object sender, System.EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            newName = tbNewName.Text;

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

            GlobalMethods.ToLog("Изменено название субъекта с '" + oldName + "' на '" + newName + "'");

            Close();
        }

        protected void btnCancel_Click(object sender, System.EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            Close();
        }
    }
}
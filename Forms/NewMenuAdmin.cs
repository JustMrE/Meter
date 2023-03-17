using System.Globalization;
using Main = Meter.MyApplicationContext;

namespace Meter.Forms
{
    public partial class NewMenuAdmin : Meter.Forms.NewMenuBase
    {

        public NewMenuAdmin()
        {
            InitializeComponent();
        }

        public override void ContextMenu()
        {
            base.ContextMenu();
            if (Main.instance.colors.main["subject"] == activeColor)
            {
                selectedButtons.Add("GoTo DB");
                selectedButtons.Add("Переименовать");
                selectedButtons.Add("Добавить новый L1");
                selectedButtons.Add("Удалить субъект");
            }
            if (activeColor == Main.instance.colors.main["прием"] || activeColor == Main.instance.colors.main["отдача"] || activeColor == Main.instance.colors.main["сальдо"])
            {
                selectedButtons.Add("Добавить новый L2");
                selectedButtons.Add("Удалить");
                if (RangeReferences.activeTable.childs.Count > 1)
                {
                    selectedButtons.Add("Удалить тип");
                }
            }
            if (Main.instance.colors.mainTitle.ContainsValue(activeColor) && RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].HasItem("аскуэ")) 
            {
                if (RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].emcosID == null)
                {
                    selectedButtons.Add("Выбрать из EMCOS");
                }
                else
                {
                    selectedButtons.Add("Изменить из EMCOS");
                    selectedButtons.Add("Удалить из EMCOS");
                }
            }
            selectedButtons.Add("Special");
        }
        public override void FormClose()
        {
            base.FormClose();
        }
        protected override void NewMenuBase_Activated(object sender, EventArgs e)
        {
            Main.instance.menu = this;
            base.NewMenuBase_Activated(sender, e);
        }
        public override void ActivateSheet(object sh)
        {
            base.ActivateSheet(sh);
        }
        public override void DeactivateSheet()
        {
            base.DeactivateSheet();
        }

        #region Buttons

        protected override void Button39_Click(object sender, EventArgs e)
        {
            base.Button39_Click(sender, e);
            GlobalMethods.ToLog("Инициализация создания нового листа...");
            Main.instance.Arhivate();
            Main.instance.references.ClearAllDB();
            string monthYear = lblMonth.Text + " " + lblYear.Text;
            DateTime date = DateTime.ParseExact(monthYear, "MMMM yyyy", GlobalMethods.culture);
            DateTime newMonthDate = date.AddMonths(1);
            lblMonth.Text = newMonthDate.ToString("MMMM", GlobalMethods.culture);
            lblYear.Text = newMonthDate.ToString("yyyy");
            Main.instance.Arhivate();
            GlobalMethods.ToLog("Лист нового месяца создан (" + newMonthDate.ToString("MMMM yyyy", GlobalMethods.culture) + ")");
            //lblMonth.Text = newMonth.ToString("MMMM", GlobalMethods.culture);
        }
        protected override void Button40_Click(object sender, EventArgs e)
        {
            base.Button40_Click(sender, e);
            GlobalMethods.ClearLogs();
        }
        protected override void Button41_Click(object sender, EventArgs e)
        {
            base.Button41_Click(sender, e);
            using (AddSubject form = new AddSubject())
            {
                form.ShowDialog();
            }
        }
        protected override void Button42_Click(object sender, EventArgs e)
        {
            base.Button42_Click(sender, e);
            using (ColorSettings colorSettings = new ColorSettings())
            {
                colorSettings.ShowDialog();
            }
        }
        protected override void btnAdmin_Click(object sender, EventArgs e)
        {
            base.btnAdmin_Click(sender, e);
            Main.instance.menu = Main.menues[0] as NewMenuBase;
            Main.instance.menu.Show();
            this.Hide();
            GlobalMethods.CalculateFormsPositions();
        }
        protected override void lblMonth_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            base.lblMonth_MouseDoubleClick(sender, e);
            using (MonthSelect form = new MonthSelect())
            {
                form.ShowDialog();
                if (!string.IsNullOrEmpty(form.selectedMonth))
                {
                    this.lblMonth.Text = form.selectedMonth;
                    month = form.selectedMonth;
                }
            }
        }
        protected override void lblYear_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            base.lblYear_MouseDoubleClick(sender, e);
            using (YearSelect form = new YearSelect(this.lblYear.Text))
            {
                form.ShowDialog();
                if (!string.IsNullOrEmpty(form.year))
                {
                    this.lblYear.Text = form.year;
                    year = form.year;
                }
            }
        }

        #endregion
    }
}

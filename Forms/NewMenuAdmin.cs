using System.Globalization;
using System.Diagnostics;
using Main = Meter.MyApplicationContext;
using Excel = Microsoft.Office.Interop.Excel;

namespace Meter.Forms
{
    public partial class NewMenuAdmin : Meter.Forms.NewMenuBase
    {
        public NewMenuAdmin()
        {
            InitializeComponent();
        }

        protected override void TextBox1_TextChanged(object sender, EventArgs e)
        {
            base.TextBox1_TextChanged(sender, e);
        }
        protected override void TextBox1_KeyUp(object sender, KeyEventArgs e)
        {
            base.TextBox1_KeyUp(sender, e);
            if (e.KeyCode == Keys.Enter)
            {
                string cmd = textBox1.Text;
                textBox1.Text = string.Empty;
                AdminCommands.RunCommand(cmd);
            }
        }
        public override void ContextMenu()
        {
            base.ContextMenu();
            if (Main.instance.colors.subColors.ContainsValue(activeColor))
            {
                CustomCellMenu.selectedButtons.Add("Удалить head");
                CustomCellMenu.selectedButtons.Add("Переименовать head");
            }
            if (Main.instance.colors.main.ContainsValue(activeColor) && Main.instance.colors.main["subject"] != activeColor)
            {
                if (RangeReferences.ActiveL1 != "план")
                {
                    if (RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].codMaketTEP == null)
                    {
                        CustomCellMenu.selectedButtons.Add("Добавить код для макетТЭП");
                    }
                    else
                    {
                        CustomCellMenu.selectedButtons.Add("Изменить код для макетТЭП");
                        CustomCellMenu.selectedButtons.Add("Удалить код для макетТЭП");
                    }

                    if (RangeReferences.activeTable.PS.childs[RangeReferences.ActiveL1].codTEP == null)
                    {
                        CustomCellMenu.selectedButtons.Add("Добавить код для ТЭП");
                    }
                    else
                    {
                        CustomCellMenu.selectedButtons.Add("Изменить код для ТЭП");
                        CustomCellMenu.selectedButtons.Add("Удалить код для ТЭП");
                    }
                    if (!RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].HasItem("счетчик"))
                    {
                        CustomCellMenu.selectedButtons.Add("Добавить по показаниям счетчика");
                    }
                    else
                    {
                        CustomCellMenu.selectedButtons.Add("Удалить по показаниям счетчика");
                    }
                }
            

            }
            if (Main.instance.colors.main["subject"] == activeColor)
            {
                CustomCellMenu.selectedButtons.Add("Переместить субъект");
                CustomCellMenu.selectedButtons.Add("GoTo DB");
                CustomCellMenu.selectedButtons.Add("Переименовать");
                CustomCellMenu.selectedButtons.Add("Добавить новый L1");
                if (RangeReferences.activeTable.DB.HasItem("план"))
                {
                    CustomCellMenu.selectedButtons.Add("Изменить код плана");
                    CustomCellMenu.selectedButtons.Add("Удалить план");
                }
                else
                {
                    CustomCellMenu.selectedButtons.Add("Добавить план");
                }
                CustomCellMenu.selectedButtons.Add("Удалить субъект");
            }
            if (activeColor == Main.instance.colors.main["прием"] || activeColor == Main.instance.colors.main["отдача"] || activeColor == Main.instance.colors.main["сальдо"])
            {
                CustomCellMenu.selectedButtons.Add("Добавить новый L2");
                CustomCellMenu.selectedButtons.Add("Удалить");
                if (RangeReferences.activeTable.PS.childs.Count > 1)
                {
                    CustomCellMenu.selectedButtons.Add("Удалить тип");
                }
            }
            if (Main.instance.colors.mainTitle.ContainsValue(activeColor) && RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].HasItem("аскуэ")) 
            {
                if (RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].emcosID == null)
                {
                    CustomCellMenu.selectedButtons.Add("Выбрать из EMCOS");
                }
                else
                {
                    CustomCellMenu.selectedButtons.Add("Изменить из EMCOS");
                    CustomCellMenu.selectedButtons.Add("Удалить из EMCOS");
                }
            }
            CustomCellMenu.selectedButtons.Add("Special");
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
        protected override void Button37_Click(object sender, EventArgs e)
        {
            base.Button37_Click(sender, e);
            Main.instance.ResumeAll();
        }
        protected override void Button38_Click(object sender, EventArgs e)
        {
            base.Button38_Click(sender, e);
            
            string monthYear = lblMonth.Text + " " + lblYear.Text;
            DateTime date = DateTime.ParseExact(monthYear, "MMMM yyyy", GlobalMethods.culture);
            DateTime newMonthDate = date.AddMonths(1);
            string newMonthYear = newMonthDate.ToString("yyyy");
            string newMonth = newMonthDate.ToString("MMMM", GlobalMethods.culture);
            string file = MeterSettings.Instance.DBDir + @"\arch\" + newMonthYear + @"\" + newMonth + @".zip";

            if (File.Exists(file))
            {
                MessageBox.Show("Лист за " + newMonthDate.ToString("MMMM yyyy", GlobalMethods.culture) + " уже существует! \nОткройте из архивов...");
                return;
            }

            MessageBox.Show("Это может занять некоторое время! \nДождитесь сообщения об окончании.");
            SplashScreen splashScreen = new SplashScreen();
            splashScreen.Show();
            GlobalMethods.ToLog("Инициализация создания нового листа...");
            splashScreen.UpdateText("Сохранение текущего листа в архив...");
            Main.instance.Arhivate();
            Main.instance.references.ClearAllDB(false, splashScreen);
            splashScreen.UpdateText("Сохранение нового листа...");
            month = newMonthDate.ToString("MMMM", GlobalMethods.culture);
            year = newMonthDate.ToString("yyyy");
            lblMonth.Text = month;
            lblYear.Text = year;
            Main.instance.Arhivate();
            GlobalMethods.ClearLogs();
            GlobalMethods.ToLog("Лист нового месяца создан (" + newMonthDate.ToString("MMMM yyyy", GlobalMethods.culture) + ")");
            splashScreen.Close();
            MessageBox.Show("Лист нового месяца создан (" + newMonthDate.ToString("MMMM yyyy", GlobalMethods.culture) + ")");
        }
        protected override void Button39_Click(object sender, EventArgs e)
        {
            base.Button39_Click(sender, e);
            try
            {
                Process.Start( new ProcessStartInfo()
                {
                    FileName = MeterSettings.Instance.LogFile,
                    UseShellExecute = true
                });
                Process.Start( new ProcessStartInfo()
                {
                    FileName = MeterSettings.Instance.ErrLogFile,
                    UseShellExecute = true
                });
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
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
            Main.instance.wsDb.Visible = Excel.XlSheetVisibility.xlSheetHidden;
        }
        protected override void lblMonth_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            base.lblMonth_MouseDoubleClick(sender, e);
            using (MonthSelect form = new MonthSelect())
            {
                form.ShowDialog();
                if (!string.IsNullOrEmpty(form.selectedMonth))
                {
                    month = form.selectedMonth;
                    this.lblMonth.Text = month;
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
                    year = form.year;
                    this.lblYear.Text = year;
                }
            }
        }

        #endregion
    }
}

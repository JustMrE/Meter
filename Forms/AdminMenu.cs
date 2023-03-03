using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Main = Meter.MyApplicationContext;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace Meter
{
    public partial class AdminMenu : MenuBase
    {
        public AdminMenu()
        {
            InitializeComponent();
            this.Load += new EventHandler(this.Menu_Load);
            this.Shown += new EventHandler(this.Menu_Show);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(Menu_Close);
        }

        public override void PreContextMenu()
        {
            base.PreContextMenu();
        }
        public override void ContextMenu()
        {
            base.ContextMenu();
            if (Main.instance.colors.main["subject"] == activeColor)
            {
                selectedButtons.Add("GoTo DB");
                // AddButtonToCommandBar("GoTo DB", GotoDB);
                selectedButtons.Add("Переименовать");
                // AddButtonToCommandBar("Переименовать", () => 
                // {
                //     using (Rename form = new Rename(RangeReferences.activeTable))
                //     {
                //         form.ShowDialog();
                //     }
                // });
                selectedButtons.Add("Добавить новый L1");
                // AddNewL1();
            }
            if (activeColor == Main.instance.colors.main["прием"] || activeColor == Main.instance.colors.main["отдача"] || activeColor == Main.instance.colors.main["сальдо"])
            {
                // AddNewL2();
                // RemoveOld();
                selectedButtons.Add("Добавить новый L2");
                selectedButtons.Add("Удалить");
            }
            
            if (Main.instance.colors.mainTitle.ContainsValue(activeColor) && RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].HasItem("аскуэ")) 
            {
                if (RangeReferences.activeTable.DB.childs[RangeReferences.ActiveL1].emcosID == null)
                {
                    // AddButtonToCommandBar("Выбрать из EMCOS", EmcosSelect);
                    selectedButtons.Add("Выбрать из EMCOS");
                }
                else
                {
                    // AddButtonToCommandBar("Изменить из EMCOS", EmcosSelect);
                    // AddButtonToCommandBar("Удалить из EMCOS", EmcosRemove);
                    selectedButtons.Add("Изменить из EMCOS");
                    selectedButtons.Add("Удалить из EMCOS");
                }
            }
            // SpecialMenuMain();
            selectedButtons.Add("Special");
        }
        protected override void RepairMenu_Click(object sender, EventArgs e)
        {
            base.RepairMenu_Click(sender, e);
        }
        protected override void btnAdmin_Click(object sender, EventArgs e)
        {
            Main.forms[0].Show();
            Main.instance.menu = (MenuBase)Main.forms[0];
            this.Hide();
            GlobalMethods.CalculateFormsPositions();
        }
        public override void ActivateSheet(object sh)
        {
            base.ActivateSheet(sh);
        }
        public override void DeactivateSheet()
        {
            base.DeactivateSheet();
        }
    }
}
﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
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
            }
            if (activeColor == Main.instance.colors.main["прием"] || activeColor == Main.instance.colors.main["отдача"] || activeColor == Main.instance.colors.main["сальдо"])
            {
                selectedButtons.Add("Добавить новый L2");
                selectedButtons.Add("Удалить");
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

        protected override void btnAdmin_Click(object sender, EventArgs e)
        {
            base.btnAdmin_Click(sender, e);
            Main.instance.menu = Main.menues[0] as NewMenuBase;
            Main.instance.menu.Show();
            this.Hide();
            GlobalMethods.CalculateFormsPositions();
        }

        protected override void Button42_Click(object sender, EventArgs e)
        {
            using (ColorSettings colorSettings = new ColorSettings())
            {
                colorSettings.ShowDialog();
            }
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
    }
}
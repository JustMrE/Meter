using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Main = Meter.MyApplicationContext;

namespace Meter
{
    public partial class ColorSettings : Form
    {
        public ColorSettings()
        {
            InitializeComponent();

            this.colorDialog1.FullOpen = true;
            //ColorsData.oldColors = new Dictionary<string, Color>(ColorsData.colors);
            //ColorsData.oldColorsForSettings1 = ColorsData.colorsForSettings.ToDictionary(n => n.Value, m => m.Value);
            ColorsData.oldColorsForSettings = new Dictionary<string, Color>(ColorsData.colorsForSettings);
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            ColorsData.UpdateColors();
        }

        private void ColorSettings_Load(object sender, EventArgs e)
        {
            Main.instance.colors.CreateAllColors();
            listBox1.Items.AddRange(Main.instance.colors.mainTitle.Keys.ToArray());
            listBox1.Items.AddRange(Main.instance.colors.extraTitle.Keys.ToArray());

            colorUp1.BackColor = Main.instance.colors.subColors["colorUp1"];
            colorUp2.BackColor = Main.instance.colors.subColors["colorUp2"];
            colorUp3.BackColor = Main.instance.colors.subColors["colorUp3"];
            colorSubject.BackColor = Main.instance.colors.main["subject"];
            button3.BackColor = Main.instance.colors.main["прием"];
            button5.BackColor = Main.instance.colors.sumColor;
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string name = (string)this.listBox1.SelectedItem;
            Color? c = Main.instance.colors.GetColor(name);
            if (c != null)
            {
                this.button4.BackColor = (Color)c;
            }
        }
        private void colorUp1_Click(object sender, EventArgs e)
        {
            Color oldColor = ((Button)sender).BackColor;
            ChangeColor(sender);
            Color newColor = ((Button)sender).BackColor;
            Main.instance.colors.allColors.Remove(oldColor);
            Main.instance.colors.allColors.Add(newColor);
            Main.instance.colors.subColors["colorUp1"] = newColor;
        }

        private void colorUp2_Click(object sender, EventArgs e)
        {
            Color oldColor = ((Button)sender).BackColor;
            ChangeColor(sender);
            Color newColor = ((Button)sender).BackColor;
            Main.instance.colors.allColors.Remove(oldColor);
            Main.instance.colors.allColors.Add(newColor);
            Main.instance.colors.subColors["colorUp2"] = newColor;
        }

        private void colorUp3_Click(object sender, EventArgs e)
        {
            Color oldColor = ((Button)sender).BackColor;
            ChangeColor(sender);
            Color newColor = ((Button)sender).BackColor;
            Main.instance.colors.allColors.Remove(oldColor);
            Main.instance.colors.allColors.Add(newColor);
            Main.instance.colors.subColors["colorUp3"] = newColor;
        }

        private void colorSubject_Click(object sender, EventArgs e)
        {
            Color oldColor = ((Button)sender).BackColor;
            ChangeColor(sender);
            Color newColor = ((Button)sender).BackColor;
            Main.instance.colors.allColors.Remove(oldColor);
            Main.instance.colors.allColors.Add(newColor);
            Main.instance.colors.main["subject"] = newColor;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Color oldColor = ((Button)sender).BackColor;
            //ChangeColor(sender);
            //Color newColor = ((Button)sender).BackColor;
            //Main.instance.colors.allColors.Remove(oldColor);
            //Main.instance.colors.allColors.Add(newColor);
            //Main.instance.colors.main["прием"] = newColor;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex >= 0)
            {
                Color oldColor = ((Button)sender).BackColor;
                ChangeColor(sender);
                Color newColor = ((Button)sender).BackColor;
                Main.instance.colors.ChangeColor((string)listBox1.SelectedItem, newColor, oldColor);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Color oldColor = ((Button)sender).BackColor;
            ChangeColor(sender);
            Color newColor = ((Button)sender).BackColor;
            Main.instance.colors.allColors.Remove(oldColor);
            Main.instance.colors.allColors.Add(newColor);
            Main.instance.colors.sumColor = newColor;
        }

        private void ChangeColor(object sender)
        {
            colorDialog1.Color = ((Button)sender).BackColor;
            if (colorDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            else
            {
                if (Main.instance.colors.IsColorFree(colorDialog1.Color))
                {
                    // установка цвета формы
                    ((Button)sender).BackColor = colorDialog1.Color;
                }
                else
                {
                    MessageBox.Show("Этот цвет уже используется!");
                    return;
                }
                
            }
        }

        private void btnLoadStandart_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            if (MessageBox.Show("Вы хотите восстановить цвета по умолчанию?", "По умолчанию", MessageBoxButtons.YesNo) == DialogResult.Yes);
            {
                ColorsData.LoadStandartColors();
                Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this, sender);
            //if (MessageBox.Show("Вы хотите отменить внесенные изменения?", "Отмена", MessageBoxButtons.YesNo) == DialogResult.Yes) ;
            {
                SaveLoader.LoadStandartColors();
                Close();
            }
        }

        private void ColorSettings_Shown(object sender, EventArgs e)
        {
            GlobalMethods.ToLog(this);
        }
    }
}

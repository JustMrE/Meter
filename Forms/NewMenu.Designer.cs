﻿namespace Meter.Forms
{
    partial class NewMenu
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        protected void InitializeComponent()
        {
            // 
            // NewMenu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(1379, 230);
            this.Name = "NewMenu";
            this.Text = "NewMenu";
            this.ResumeLayout(false);
            this.PerformLayout();
            //
            //button 2
            //
            this.button2.Visible = true;
            this.button2.Text = "EMCOS";
            //
            //button 3
            //
            this.button3.Visible = true;
            this.button3.Text = "записать счетчики";
            //
            //button 4
            //
            this.button4.Visible = true;
            this.button4.Text = "старые счетчики (временно)";
            //
            //button 5
            //
            this.button5.Visible = true;
            this.button5.Text = "записать ТЭП";
            //
            //button 6
            //
            this.button6.Visible = true;
            this.button6.Text = "записать планы";
            
        }

        #endregion
    }
}
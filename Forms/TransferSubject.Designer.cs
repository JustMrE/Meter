namespace Meter.Forms
{
    partial class TransferSubject
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.RBleft = new System.Windows.Forms.RadioButton();
            this.RBright = new System.Windows.Forms.RadioButton();
            this.RBzone = new System.Windows.Forms.RadioButton();
            this.flpAll = new System.Windows.Forms.FlowLayoutPanel();
            this.FlowLayoutPanel3 = new System.Windows.Forms.FlowLayoutPanel();
            this.ComboBox11 = new System.Windows.Forms.ComboBox();
            this.ComboBox12 = new System.Windows.Forms.ComboBox();
            this.ComboBox13 = new System.Windows.Forms.ComboBox();
            this.FlowLayoutPanel20 = new System.Windows.Forms.FlowLayoutPanel();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.flpAll.SuspendLayout();
            this.FlowLayoutPanel3.SuspendLayout();
            this.FlowLayoutPanel20.SuspendLayout();
            this.SuspendLayout();
            //
            // RadioButton1
            //
            this.RBleft.AutoSize =  true;
            this.RBleft.Text =  "Влево";
            this.RBleft.Location = new System.Drawing.Point(3,135);
            this.RBleft.Size = new System.Drawing.Size(103,19);
            this.RBleft.TabIndex = 11;
            this.RBleft.CheckedChanged += new System.EventHandler(this.RadioButton1_CheckedChanged);
            //
            // RadioButton12
            //
            this.RBright.AutoSize =  true;
            this.RBright.Text =  "Вправо";
            this.RBright.Location = new System.Drawing.Point(3,160);
            this.RBright.Size = new System.Drawing.Size(103,19);
            this.RBright.TabIndex = 12;
            this.RBright.CheckedChanged += new System.EventHandler(this.RadioButton2_CheckedChanged);
            //
            // RadioButton13
            //
            this.RBzone.AutoSize =  true;
            this.RBzone.Text =  "В область";
            this.RBzone.Location = new System.Drawing.Point(3,185);
            this.RBzone.Size = new System.Drawing.Size(103,19);
            this.RBzone.TabIndex = 13;
            this.RBzone.CheckedChanged += new System.EventHandler(this.RadioButton3_CheckedChanged);
            //
            // flpAll
            //
            this.flpAll.Controls.Add(this.RBleft);
            this.flpAll.Controls.Add(this.RBright);
            this.flpAll.Controls.Add(this.RBzone);
            this.flpAll.Controls.Add(this.FlowLayoutPanel3);
            this.flpAll.Controls.Add(this.FlowLayoutPanel20);
            this.flpAll.AutoSize =  true;
            this.flpAll.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.flpAll.FlowDirection = FlowDirection.TopDown;
            this.flpAll.Text =  "FlowLayoutPanel0";
            this.flpAll.Location = new System.Drawing.Point(8,8);
            this.flpAll.Size = new System.Drawing.Size(316,134);
            //
            // FlowLayoutPanel3
            //
            this.FlowLayoutPanel3.Controls.Add(this.ComboBox11);
            this.FlowLayoutPanel3.Controls.Add(this.ComboBox12);
            this.FlowLayoutPanel3.Controls.Add(this.ComboBox13);
            this.FlowLayoutPanel3.AutoSize =  true;
            this.FlowLayoutPanel3.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.FlowLayoutPanel3.FlowDirection = FlowDirection.TopDown;
            this.FlowLayoutPanel3.Text =  "FlowLayoutPanel3";
            this.FlowLayoutPanel3.BackColor = Color.Transparent;
            this.FlowLayoutPanel3.Location = new System.Drawing.Point(3,3);
            this.FlowLayoutPanel3.Size = new System.Drawing.Size(308,89);
            this.FlowLayoutPanel3.TabIndex = 3;
            //
            // ComboBox11
            //
            this.ComboBox11.DropDownWidth = 300;
            this.ComboBox11.ItemHeight = 15;
            this.ComboBox11.Text =  "ComboBox11";
            this.ComboBox11.Location = new System.Drawing.Point(3,3);
            this.ComboBox11.Size = new System.Drawing.Size(300,23);
            this.ComboBox11.TabIndex = 11;
            this.ComboBox11.TextChanged += new System.EventHandler(this.ComboBox11_TextChanged);
            //
            // ComboBox12
            //
            this.ComboBox12.DropDownWidth = 300;
            this.ComboBox12.ItemHeight = 15;
            this.ComboBox12.Text =  "ComboBox12";
            this.ComboBox12.Location = new System.Drawing.Point(3,32);
            this.ComboBox12.Size = new System.Drawing.Size(300,23);
            this.ComboBox12.TabIndex = 12;
            this.ComboBox12.TextChanged += new System.EventHandler(this.ComboBox12_TextChanged);
            //
            // ComboBox13
            //
            this.ComboBox13.DropDownWidth = 300;
            this.ComboBox13.ItemHeight = 15;
            this.ComboBox13.Text =  "ComboBox13";
            this.ComboBox13.Location = new System.Drawing.Point(3,61);
            this.ComboBox13.Size = new System.Drawing.Size(300,23);
            this.ComboBox13.TabIndex = 13;
            this.ComboBox13.TextChanged += new System.EventHandler(this.ComboBox13_TextChanged);
            //
            // FlowLayoutPanel20
            //
            this.FlowLayoutPanel20.Controls.Add(this.btnOk);
            this.FlowLayoutPanel20.Controls.Add(this.btnCancel);
            this.FlowLayoutPanel20.AutoSize =  true;
            this.FlowLayoutPanel20.Text =  "FlowLayoutPanel20";
            this.FlowLayoutPanel20.BackColor = Color.Transparent;
            this.FlowLayoutPanel20.Location = new System.Drawing.Point(147,98);
            this.FlowLayoutPanel20.Size = new System.Drawing.Size(164,31);
            this.FlowLayoutPanel20.TabIndex = 20;
            this.FlowLayoutPanel20.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            //
            // btnOk
            //
            this.btnOk.BackColor = Color.Transparent;
            this.btnOk.Text =  "Ok";
            this.btnOk.Location = new System.Drawing.Point(3,3);
            this.btnOk.TabIndex = 21;
            this.btnOk.Click += new System.EventHandler(this.BtnOk_Click);
            //
            // btnCancel
            //
            this.btnCancel.BackColor = Color.Transparent;
            this.btnCancel.Text =  "Cancel";
            this.btnCancel.Location = new System.Drawing.Point(84,3);
            this.btnCancel.TabIndex = 22;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
         //
         // form
         //
            this.AutoSize =  true;
            this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Size = new System.Drawing.Size(355,196);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text =  "Переместить субъект";
            this.Shown += new System.EventHandler(this.TransferSubject_Shown);
            this.Controls.Add(this.flpAll);
            this.flpAll.ResumeLayout(false);
            this.FlowLayoutPanel3.ResumeLayout(false);
            this.FlowLayoutPanel20.ResumeLayout(false);
            this.ResumeLayout(false);
        } 

        #endregion 

        private System.Windows.Forms.RadioButton RBleft;
        private System.Windows.Forms.RadioButton RBright;
        private System.Windows.Forms.RadioButton RBzone;
        private System.Windows.Forms.FlowLayoutPanel flpAll;
        private System.Windows.Forms.FlowLayoutPanel FlowLayoutPanel3;
        private System.Windows.Forms.ComboBox ComboBox11;
        private System.Windows.Forms.ComboBox ComboBox12;
        private System.Windows.Forms.ComboBox ComboBox13;
        private System.Windows.Forms.FlowLayoutPanel FlowLayoutPanel20;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
    }
}


namespace Meter.Forms
{
    partial class AddSubject
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
            this.flpAll = new System.Windows.Forms.FlowLayoutPanel();
            this.FlowLayoutPanel3 = new System.Windows.Forms.FlowLayoutPanel();
            this.ComboBox11 = new System.Windows.Forms.ComboBox();
            this.ComboBox12 = new System.Windows.Forms.ComboBox();
            this.ComboBox13 = new System.Windows.Forms.ComboBox();
            this.FlowLayoutPanel4 = new System.Windows.Forms.FlowLayoutPanel();
            this.Label14 = new System.Windows.Forms.Label();
            this.TextBox15 = new System.Windows.Forms.TextBox();
            this.FlowLayoutPanel16 = new System.Windows.Forms.FlowLayoutPanel();
            this.CheckBox17 = new System.Windows.Forms.CheckBox();
            this.CheckBox18 = new System.Windows.Forms.CheckBox();
            this.CheckBox19 = new System.Windows.Forms.CheckBox();
            this.FlowLayoutPanel20 = new System.Windows.Forms.FlowLayoutPanel();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.flpAll.SuspendLayout();
            this.FlowLayoutPanel3.SuspendLayout();
            this.FlowLayoutPanel4.SuspendLayout();
            this.FlowLayoutPanel16.SuspendLayout();
            this.FlowLayoutPanel20.SuspendLayout();
            this.SuspendLayout();
            // 
            // flpAll
            // 
            this.flpAll.AutoSize = true;
            this.flpAll.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.flpAll.Controls.Add(this.FlowLayoutPanel3);
            this.flpAll.Controls.Add(this.FlowLayoutPanel4);
            this.flpAll.Controls.Add(this.FlowLayoutPanel16);
            this.flpAll.Controls.Add(this.FlowLayoutPanel20);
            this.flpAll.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flpAll.Location = new System.Drawing.Point(8, 8);
            this.flpAll.Name = "flpAll";
            this.flpAll.Size = new System.Drawing.Size(429, 194);
            this.flpAll.TabIndex = 0;
            this.flpAll.Text = "FlowLayoutPanel0";
            // 
            // FlowLayoutPanel3
            // 
            this.FlowLayoutPanel3.AutoSize = true;
            this.FlowLayoutPanel3.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.FlowLayoutPanel3.BackColor = System.Drawing.Color.Transparent;
            this.FlowLayoutPanel3.Controls.Add(this.ComboBox11);
            this.FlowLayoutPanel3.Controls.Add(this.ComboBox12);
            this.FlowLayoutPanel3.Controls.Add(this.ComboBox13);
            this.FlowLayoutPanel3.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.FlowLayoutPanel3.Location = new System.Drawing.Point(120, 3);
            this.FlowLayoutPanel3.Margin = new System.Windows.Forms.Padding(120, 3, 3, 3);
            this.FlowLayoutPanel3.Name = "FlowLayoutPanel3";
            this.FlowLayoutPanel3.Size = new System.Drawing.Size(306, 87);
            this.FlowLayoutPanel3.TabIndex = 3;
            this.FlowLayoutPanel3.Text = "FlowLayoutPanel3";
            // 
            // ComboBox11
            // 
            this.ComboBox11.DropDownWidth = 300;
            this.ComboBox11.ItemHeight = 15;
            this.ComboBox11.Location = new System.Drawing.Point(3, 3);
            this.ComboBox11.Name = "ComboBox11";
            this.ComboBox11.Size = new System.Drawing.Size(300, 23);
            this.ComboBox11.TabIndex = 11;
            this.ComboBox11.TextChanged += new System.EventHandler(this.ComboBox11_TextChanged);
            // 
            // ComboBox12
            // 
            this.ComboBox12.DropDownWidth = 300;
            this.ComboBox12.ItemHeight = 15;
            this.ComboBox12.Location = new System.Drawing.Point(3, 32);
            this.ComboBox12.Name = "ComboBox12";
            this.ComboBox12.Size = new System.Drawing.Size(300, 23);
            this.ComboBox12.TabIndex = 12;
            this.ComboBox12.Visible = false;
            this.ComboBox12.TextChanged += new System.EventHandler(this.ComboBox12_TextChanged);
            // 
            // ComboBox13
            // 
            this.ComboBox13.DropDownWidth = 300;
            this.ComboBox13.ItemHeight = 15;
            this.ComboBox13.Location = new System.Drawing.Point(3, 61);
            this.ComboBox13.Name = "ComboBox13";
            this.ComboBox13.Size = new System.Drawing.Size(300, 23);
            this.ComboBox13.TabIndex = 13;
            this.ComboBox13.Visible = false;
            this.ComboBox13.TextChanged += new System.EventHandler(this.ComboBox13_TextChanged);
            // 
            // FlowLayoutPanel4
            // 
            this.FlowLayoutPanel4.AutoSize = true;
            this.FlowLayoutPanel4.BackColor = System.Drawing.Color.Transparent;
            this.FlowLayoutPanel4.Controls.Add(this.Label14);
            this.FlowLayoutPanel4.Controls.Add(this.TextBox15);
            this.FlowLayoutPanel4.Location = new System.Drawing.Point(3, 96);
            this.FlowLayoutPanel4.Name = "FlowLayoutPanel4";
            this.FlowLayoutPanel4.Size = new System.Drawing.Size(423, 29);
            this.FlowLayoutPanel4.TabIndex = 4;
            this.FlowLayoutPanel4.Text = "FlowLayoutPanel4";
            // 
            // Label14
            // 
            this.Label14.AutoSize = true;
            this.Label14.BackColor = System.Drawing.Color.Transparent;
            this.Label14.Location = new System.Drawing.Point(3, 0);
            this.Label14.Name = "Label14";
            this.Label14.Size = new System.Drawing.Size(111, 15);
            this.Label14.TabIndex = 14;
            this.Label14.Text = "Название субъекта";
            // 
            // TextBox15
            // 
            this.TextBox15.Location = new System.Drawing.Point(120, 3);
            this.TextBox15.Name = "TextBox15";
            this.TextBox15.Size = new System.Drawing.Size(300, 23);
            this.TextBox15.TabIndex = 15;
            // 
            // FlowLayoutPanel16
            // 
            this.FlowLayoutPanel16.AutoSize = true;
            this.FlowLayoutPanel16.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.FlowLayoutPanel16.BackColor = System.Drawing.Color.Transparent;
            this.FlowLayoutPanel16.Controls.Add(this.CheckBox17);
            this.FlowLayoutPanel16.Controls.Add(this.CheckBox18);
            this.FlowLayoutPanel16.Controls.Add(this.CheckBox19);
            this.FlowLayoutPanel16.Location = new System.Drawing.Point(120, 131);
            this.FlowLayoutPanel16.Margin = new System.Windows.Forms.Padding(120, 3, 3, 3);
            this.FlowLayoutPanel16.Name = "FlowLayoutPanel16";
            this.FlowLayoutPanel16.Size = new System.Drawing.Size(207, 25);
            this.FlowLayoutPanel16.TabIndex = 16;
            this.FlowLayoutPanel16.Text = "FlowLayoutPanel16";
            // 
            // CheckBox17
            // 
            this.CheckBox17.AutoSize = true;
            this.CheckBox17.BackColor = System.Drawing.Color.Transparent;
            this.CheckBox17.Location = new System.Drawing.Point(3, 3);
            this.CheckBox17.Name = "CheckBox17";
            this.CheckBox17.Size = new System.Drawing.Size(62, 19);
            this.CheckBox17.TabIndex = 17;
            this.CheckBox17.Text = "прием";
            this.CheckBox17.UseVisualStyleBackColor = false;
            this.CheckBox17.CheckedChanged += new System.EventHandler(this.CheckBox_CheckedChanged);
            // 
            // CheckBox18
            // 
            this.CheckBox18.AutoSize = true;
            this.CheckBox18.BackColor = System.Drawing.Color.Transparent;
            this.CheckBox18.Location = new System.Drawing.Point(71, 3);
            this.CheckBox18.Name = "CheckBox18";
            this.CheckBox18.Size = new System.Drawing.Size(63, 19);
            this.CheckBox18.TabIndex = 18;
            this.CheckBox18.Text = "отдача";
            this.CheckBox18.UseVisualStyleBackColor = false;
            this.CheckBox18.CheckedChanged += new System.EventHandler(this.CheckBox_CheckedChanged);
            // 
            // CheckBox19
            // 
            this.CheckBox19.AutoSize = true;
            this.CheckBox19.BackColor = System.Drawing.Color.Transparent;
            this.CheckBox19.Location = new System.Drawing.Point(140, 3);
            this.CheckBox19.Name = "CheckBox19";
            this.CheckBox19.Size = new System.Drawing.Size(64, 19);
            this.CheckBox19.TabIndex = 19;
            this.CheckBox19.Text = "сальдо";
            this.CheckBox19.UseVisualStyleBackColor = false;
            this.CheckBox19.CheckedChanged += new System.EventHandler(this.CheckBox_CheckedChanged);
            // 
            // FlowLayoutPanel20
            // 
            this.FlowLayoutPanel20.AutoSize = true;
            this.FlowLayoutPanel20.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.FlowLayoutPanel20.BackColor = System.Drawing.Color.Transparent;
            this.FlowLayoutPanel20.Controls.Add(this.btnOk);
            this.FlowLayoutPanel20.Controls.Add(this.btnCancel);
            this.FlowLayoutPanel20.Location = new System.Drawing.Point(145, 162);
            this.FlowLayoutPanel20.Margin = new System.Windows.Forms.Padding(145, 3, 3, 3);
            this.FlowLayoutPanel20.Name = "FlowLayoutPanel20";
            this.FlowLayoutPanel20.Size = new System.Drawing.Size(162, 29);
            this.FlowLayoutPanel20.TabIndex = 20;
            this.FlowLayoutPanel20.Text = "FlowLayoutPanel20";
            // 
            // btnOk
            // 
            this.btnOk.BackColor = System.Drawing.Color.Transparent;
            this.btnOk.Location = new System.Drawing.Point(3, 3);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 21;
            this.btnOk.Text = "Ok";
            this.btnOk.UseVisualStyleBackColor = false;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.Location = new System.Drawing.Point(84, 3);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 22;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // AddSubject
            // 
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(444, 215);
            this.Controls.Add(this.flpAll);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Name = "AddSubject";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Добавить субъект";
            this.Shown += new System.EventHandler(this.AddSubject_Shown);
            this.flpAll.ResumeLayout(false);
            this.flpAll.PerformLayout();
            this.FlowLayoutPanel3.ResumeLayout(false);
            this.FlowLayoutPanel4.ResumeLayout(false);
            this.FlowLayoutPanel4.PerformLayout();
            this.FlowLayoutPanel16.ResumeLayout(false);
            this.FlowLayoutPanel16.PerformLayout();
            this.FlowLayoutPanel20.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        } 

        #endregion 

        private System.Windows.Forms.FlowLayoutPanel flpAll;
        private System.Windows.Forms.FlowLayoutPanel FlowLayoutPanel3;
        private System.Windows.Forms.FlowLayoutPanel FlowLayoutPanel4;
        private System.Windows.Forms.ComboBox ComboBox11;
        private System.Windows.Forms.ComboBox ComboBox12;
        private System.Windows.Forms.ComboBox ComboBox13;
        private System.Windows.Forms.Label Label14;
        private System.Windows.Forms.TextBox TextBox15;
        private System.Windows.Forms.FlowLayoutPanel FlowLayoutPanel16;
        private System.Windows.Forms.CheckBox CheckBox17;
        private System.Windows.Forms.CheckBox CheckBox18;
        private System.Windows.Forms.CheckBox CheckBox19;
        private System.Windows.Forms.FlowLayoutPanel FlowLayoutPanel20;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
    }
}


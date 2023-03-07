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
            this.FlowLayoutPanel4 = new System.Windows.Forms.FlowLayoutPanel();
            this.ComboBox11 = new System.Windows.Forms.ComboBox();
            this.ComboBox12 = new System.Windows.Forms.ComboBox();
            this.ComboBox13 = new System.Windows.Forms.ComboBox();
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
            this.flpAll.Controls.Add(this.FlowLayoutPanel3);
            this.flpAll.Controls.Add(this.FlowLayoutPanel4);
            this.flpAll.Controls.Add(this.FlowLayoutPanel16);
            this.flpAll.Controls.Add(this.FlowLayoutPanel20);
            this.flpAll.AutoSize =  true;
            this.flpAll.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.flpAll.FlowDirection = FlowDirection.TopDown;
            this.flpAll.Text =  "FlowLayoutPanel0";
            this.flpAll.Location = new System.Drawing.Point(8,8);
            this.flpAll.Size = new System.Drawing.Size(433,204);
            //
            // FlowLayoutPanel3
            //
            this.FlowLayoutPanel3.Controls.Add(this.ComboBox11);
            this.FlowLayoutPanel3.Controls.Add(this.ComboBox12);
            this.FlowLayoutPanel3.Controls.Add(this.ComboBox13);
            this.FlowLayoutPanel3.AutoSize =  true;
            this.FlowLayoutPanel3.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.FlowLayoutPanel3.FlowDirection = FlowDirection.TopDown;
            this.FlowLayoutPanel3.Margin = new Padding(120, 3, 3, 3);
            this.FlowLayoutPanel3.Text =  "FlowLayoutPanel3";
            this.FlowLayoutPanel3.BackColor = Color.Transparent;
            this.FlowLayoutPanel3.Location = new System.Drawing.Point(120,3);
            this.FlowLayoutPanel3.Size = new System.Drawing.Size(308,89);
            this.FlowLayoutPanel3.TabIndex = 3;
            //
            // FlowLayoutPanel4
            //
            this.FlowLayoutPanel4.Controls.Add(this.Label14);
            this.FlowLayoutPanel4.Controls.Add(this.TextBox15);
            this.FlowLayoutPanel4.AutoSize =  true;
            this.FlowLayoutPanel4.Text =  "FlowLayoutPanel4";
            this.FlowLayoutPanel4.BackColor = Color.Transparent;
            this.FlowLayoutPanel4.Location = new System.Drawing.Point(3,98);
            this.FlowLayoutPanel4.Size = new System.Drawing.Size(425,31);
            this.FlowLayoutPanel4.TabIndex = 4;
            //
            // ComboBox11
            //
            this.ComboBox11.DropDownWidth = 300;
            this.ComboBox11.ItemHeight = 15;
            this.ComboBox11.Location = new System.Drawing.Point(3,3);
            this.ComboBox11.Size = new System.Drawing.Size(300,23);
            this.ComboBox11.TabIndex = 11;
            this.ComboBox11.TextChanged += new EventHandler(ComboBox11_TextChanged);
            //
            // ComboBox12
            //
            this.ComboBox12.DropDownWidth = 300;
            this.ComboBox12.ItemHeight = 15;
            this.ComboBox12.Location = new System.Drawing.Point(3,32);
            this.ComboBox12.Size = new System.Drawing.Size(300,23);
            this.ComboBox12.TabIndex = 12;
            this.ComboBox12.Visible = false;
            this.ComboBox12.TextChanged += new EventHandler(ComboBox12_TextChanged);
            //
            // ComboBox13
            //
            this.ComboBox13.DropDownWidth = 300;
            this.ComboBox13.ItemHeight = 15;
            this.ComboBox13.Location = new System.Drawing.Point(3,61);
            this.ComboBox13.Size = new System.Drawing.Size(300,23);
            this.ComboBox13.TabIndex = 13;
            this.ComboBox13.Visible = false;
            this.ComboBox13.TextChanged += new EventHandler(ComboBox13_TextChanged);
            //
            // Label14
            //
            this.Label14.AutoSize =  true;
            this.FlowLayoutPanel20.Margin = new Padding(3, 7, 3, 3);
            this.Label14.Text =  "Название субъекта";
            this.Label14.BackColor = Color.Transparent;
            this.Label14.Location = new System.Drawing.Point(3,0);
            this.Label14.Size = new System.Drawing.Size(111,15);
            this.Label14.TabIndex = 14;
            //
            // TextBox15
            //
            this.TextBox15.Location = new System.Drawing.Point(120,3);
            this.TextBox15.Size = new System.Drawing.Size(300,23);
            this.TextBox15.TabIndex = 15;
            //
            // FlowLayoutPanel16
            //
            this.FlowLayoutPanel16.Controls.Add(this.CheckBox17);
            this.FlowLayoutPanel16.Controls.Add(this.CheckBox18);
            this.FlowLayoutPanel16.Controls.Add(this.CheckBox19);
            this.FlowLayoutPanel16.AutoSize =  true;
            this.FlowLayoutPanel16.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.FlowLayoutPanel16.Margin = new Padding(120, 3, 3, 3);
            this.FlowLayoutPanel16.Text =  "FlowLayoutPanel16";
            this.FlowLayoutPanel16.BackColor = Color.Transparent;
            this.FlowLayoutPanel16.Location = new System.Drawing.Point(120,135);
            this.FlowLayoutPanel16.Size = new System.Drawing.Size(209,27);
            this.FlowLayoutPanel16.TabIndex = 16;
            //
            // CheckBox17
            //
            this.CheckBox17.Checked =  false;
            this.CheckBox17.AutoSize =  true;
            this.CheckBox17.BackColor = Color.Transparent;
            this.CheckBox17.Text =  "прием";
            this.CheckBox17.Location = new System.Drawing.Point(3,3);
            this.CheckBox17.Size = new System.Drawing.Size(62,19);
            this.CheckBox17.TabIndex = 17;
            //
            // CheckBox18
            //
            this.CheckBox18.Checked =  false;
            this.CheckBox18.AutoSize =  true;
            this.CheckBox18.BackColor = Color.Transparent;
            this.CheckBox18.Text =  "отдача";
            this.CheckBox18.Location = new System.Drawing.Point(71,3);
            this.CheckBox18.Size = new System.Drawing.Size(63,19);
            this.CheckBox18.TabIndex = 18;
            //
            // CheckBox19
            //
            this.CheckBox19.Checked =  false;
            this.CheckBox19.AutoSize =  true;
            this.CheckBox19.BackColor = Color.Transparent;
            this.CheckBox19.Text =  "сальдо";
            this.CheckBox19.Location = new System.Drawing.Point(140,3);
            this.CheckBox19.Size = new System.Drawing.Size(64,19);
            this.CheckBox19.TabIndex = 19;
            //
            // FlowLayoutPanel20
            //
            this.FlowLayoutPanel20.Controls.Add(this.btnOk);
            this.FlowLayoutPanel20.Controls.Add(this.btnCancel);
            this.FlowLayoutPanel20.AutoSize =  true;
            this.FlowLayoutPanel20.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.FlowLayoutPanel20.Margin = new Padding(145, 3, 3, 3);
            this.FlowLayoutPanel20.Text =  "FlowLayoutPanel20";
            this.FlowLayoutPanel20.BackColor = Color.Transparent;
            this.FlowLayoutPanel20.Location = new System.Drawing.Point(145,168);
            this.FlowLayoutPanel20.Size = new System.Drawing.Size(164,31);
            this.FlowLayoutPanel20.TabIndex = 20;
            //
            // btnOk
            //
            this.btnOk.BackColor = Color.Transparent;
            this.btnOk.Text =  "Ok";
            this.btnOk.Location = new System.Drawing.Point(3,3);
            this.btnOk.TabIndex = 21;
            this.btnOk.Click += new EventHandler(btnOk_Click);
            //
            // btnCancel
            //
            this.btnCancel.BackColor = Color.Transparent;
            this.btnCancel.Text =  "Cancel";
            this.btnCancel.Location = new System.Drawing.Point(84,3);
            this.btnCancel.TabIndex = 22;
            this.btnCancel.Click += new EventHandler(btnCancel_Click);
         //
         // form
         //
            this.AutoSize =  true;
            this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Size = new System.Drawing.Size(464,258);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text =  "Добавить субъект";
            this.Controls.Add(this.flpAll);
            this.flpAll.ResumeLayout(false);
            this.FlowLayoutPanel3.ResumeLayout(false);
            this.FlowLayoutPanel4.ResumeLayout(false);
            this.FlowLayoutPanel16.ResumeLayout(false);
            this.FlowLayoutPanel20.ResumeLayout(false);
            this.ResumeLayout(false);
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


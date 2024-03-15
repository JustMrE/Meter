namespace Meter.Forms
{
    partial class FormulaEditor
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
        private void InitializeComponent()
        {
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
            this.tbSearch = new System.Windows.Forms.TextBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.trash = new System.Windows.Forms.FlowLayoutPanel();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnConst = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.cbSearchFormula = new System.Windows.Forms.CheckBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnLoad = new System.Windows.Forms.Button();
            this.GroupBox18 = new GroupBox();
            this.CheckBox19 = new CheckBox();
            this.MonthCalendar20 = new MonthCalendar();
            this.SuspendLayout();
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.AutoScroll = true;
            this.flowLayoutPanel1.BackColor = System.Drawing.SystemColors.Info;
            this.flowLayoutPanel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(12, 12);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(1044,287);
            this.flowLayoutPanel1.TabIndex = 0;
            // 
            // listView1
            // 
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.listView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.listView1.Location = new System.Drawing.Point(12, 334);
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(660,160);
            this.listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.listView1.TabIndex = 3;
            this.listView1.TileSize = new System.Drawing.Size(10, 10);
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.listView1_MouseDown);
            this.listView1.Click += new System.EventHandler(this.listView1_Click);
            // 
            // tbSearch
            // 
            this.tbSearch.Location = new System.Drawing.Point(12, 305);
            this.tbSearch.Name = "tbSearch";
            this.tbSearch.PlaceholderText = "Поиск...";
            this.tbSearch.Size = new System.Drawing.Size(492, 23);
            this.tbSearch.TabIndex = 4;
            this.tbSearch.TextChanged += new System.EventHandler(this.tbSearch_TextChanged);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(518,306);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(151,19);
            this.checkBox1.TabIndex = 5;
            this.checkBox1.Text = "Учитывать регистр";
            this.checkBox1.UseVisualStyleBackColor = true;
            //
            // cbSearchFormula
            //
            this.cbSearchFormula.AutoSize =  true;
            this.cbSearchFormula.Name = "cbSearchFormula";
            this.cbSearchFormula.Text =  "Поиск в формуле";
            this.cbSearchFormula.Location = new System.Drawing.Point(676,306);
            this.cbSearchFormula.Size = new System.Drawing.Size(124,19);
            this.cbSearchFormula.TabIndex = 15;
            this.cbSearchFormula.UseVisualStyleBackColor = true;
            this.cbSearchFormula.CheckedChanged += new System.EventHandler(this.cbSearchFormula_CheckedChanged);
            // 
            // button1
            // 
            this.button1.Name = "button1";
            this.button1.Location = new System.Drawing.Point(686,333);
            this.button1.Size = new System.Drawing.Size(37,23);
            this.button1.TabIndex = 6;
            this.button1.Text = "+";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.control1_MouseDown);
            // 
            // button2
            // 
            this.button2.Name = "button2";
            this.button2.Location = new System.Drawing.Point(686,358);
            this.button2.Size = new System.Drawing.Size(37,23);
            this.button2.TabIndex = 7;
            this.button2.Text = "*";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.MouseDown += new System.Windows.Forms.MouseEventHandler(this.control1_MouseDown);
            // 
            // button3
            // 
            this.button3.Name = "button3";
            this.button3.Location = new System.Drawing.Point(729,333);
            this.button3.Size = new System.Drawing.Size(37,23);
            this.button3.TabIndex = 8;
            this.button3.Text = "-";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.MouseDown += new System.Windows.Forms.MouseEventHandler(this.control1_MouseDown);
            // 
            // button4
            // 
            this.button4.Name = "button4";
            this.button4.Location = new System.Drawing.Point(729,358);
            this.button4.Size = new System.Drawing.Size(37,23);
            this.button4.TabIndex = 9;
            this.button4.Text = "/";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.MouseDown += new System.Windows.Forms.MouseEventHandler(this.control1_MouseDown);
            // 
            // button5
            // 
            this.button5.ForeColor = System.Drawing.SystemColors.ControlText;
            this.button5.Name = "button5";
            this.button5.Location = new System.Drawing.Point(686,383);
            this.button5.Size = new System.Drawing.Size(37,23);
            this.button5.TabIndex = 10;
            this.button5.Text = "(";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.MouseDown += new System.Windows.Forms.MouseEventHandler(this.control1_MouseDown);
            // 
            // button6
            // 
            this.button6.Name = "button6";
            this.button6.Location = new System.Drawing.Point(729,383);
            this.button6.Size = new System.Drawing.Size(37,23);
            this.button6.TabIndex = 11;
            this.button6.Text = ")";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.MouseDown += new System.Windows.Forms.MouseEventHandler(this.control1_MouseDown);
            // 
            // trash
            // 
            this.trash.AllowDrop = true;
            this.trash.BackgroundImage = global::Meter.Properties.Resources.trash;
            this.trash.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.trash.Location = new System.Drawing.Point(5, 259);
            this.trash.Name = "trash";
            this.trash.Size = new System.Drawing.Size(35, 40);
            this.trash.TabIndex = 13;
            this.trash.Visible = false;
            this.trash.DragDrop += new System.Windows.Forms.DragEventHandler(this.trash_DragDrop);
            this.trash.DragEnter += new System.Windows.Forms.DragEventHandler(this.trash_DragEnter);
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(314, 506);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 14;
            this.btnOk.Text = "Ok";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnClear
            // 
            this.btnClear.Name = "btnClear";
            this.btnClear.Location = new System.Drawing.Point(686,437);
            this.btnClear.Size = new System.Drawing.Size(80,23);
            this.btnClear.TabIndex = 17;
            this.btnClear.Text = "Очистить";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnConst
            // 
            this.btnConst.Name = "btnConst";
            this.btnConst.Location = new System.Drawing.Point(686,409);
            this.btnConst.Size = new System.Drawing.Size(80,23);
            this.btnConst.TabIndex = 16;
            this.btnConst.Text = "Константа";
            this.btnConst.UseVisualStyleBackColor = true;
            this.btnConst.MouseDown += new System.Windows.Forms.MouseEventHandler(this.control1_MouseDown);
            //
            // btnSave
            //
            this.btnSave.Text =  "Экспорт";
            this.btnSave.Location = new System.Drawing.Point(772,409);
            this.btnSave.TabIndex = 16;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            //
            // btnLoad
            //
            this.btnLoad.Text =  "Импорт";
            this.btnLoad.Location = new System.Drawing.Point(772,437);
            this.btnLoad.TabIndex = 17;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(405, 506);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 15;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            //
            // GroupBox18
            //
            this.GroupBox18.Controls.Add(this.CheckBox19);
            this.GroupBox18.Controls.Add(this.MonthCalendar20);
            this.GroupBox18.Location = new System.Drawing.Point(856,304);
            this.GroupBox18.Size = new System.Drawing.Size(200,220);
            this.GroupBox18.TabIndex = 18;
            //
            // CheckBox19
            //
            this.CheckBox19.AutoSize =  true;
            this.CheckBox19.Text =  "Проверка";
            this.CheckBox19.Location = new System.Drawing.Point(12,20);
            this.CheckBox19.Size = new System.Drawing.Size(80,19);
            this.CheckBox19.TabIndex = 19;
            this.CheckBox19.CheckedChanged += new EventHandler(CheckBox19_CheckedChanged);
            //
            // MonthCalendar20
            //
            this.MonthCalendar20.MaxSelectionCount = 1;
            this.MonthCalendar20.Size = new System.Drawing.Size(164,162);
            this.MonthCalendar20.Text =  "MonthCalendar20";
            this.MonthCalendar20.Location = new System.Drawing.Point(12,48);
            this.MonthCalendar20.TabIndex = 20;
            this.MonthCalendar20.Visible = false;
            this.MonthCalendar20.DateChanged += new DateRangeEventHandler(MonthCalendar20_DateChanged);
            this.MonthCalendar20.DateSelected += new DateRangeEventHandler(MonthCalendar20_DateSelected);
            // 
            // FormulaEditor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1084,541);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.trash);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.tbSearch);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.flowLayoutPanel1);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnConst);
            this.Controls.Add(this.cbSearchFormula);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.GroupBox18);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Name = "FormulaEditor";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormulaEditor";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormulaEditor_Closing);
            this.Load += new System.EventHandler(this.FormulaEditor_Load);
            this.Shown += new System.EventHandler(this.FormulaEditor_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private FlowLayoutPanel flowLayoutPanel1;
        private ListView listView1;
        private TextBox tbSearch;
        private CheckBox checkBox1;
        private ColumnHeader columnHeader1;
        private Button button1;
        private Button button2;
        private Button button3;
        private Button button4;
        private Button button5;
        private Button button6;
        private FlowLayoutPanel trash;
        private Button btnOk;
        private Button btnCancel;
        private Button btnClear;
        private Button btnConst;
        private CheckBox cbSearchFormula;
        private Button btnSave;
        private Button btnLoad;

        private GroupBox GroupBox18;
        private CheckBox CheckBox19;
        private MonthCalendar MonthCalendar20;
    }
}
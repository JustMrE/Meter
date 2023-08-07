namespace Meter.Forms
{
    partial class AllFormulas
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
            // this.listBox0 = new System.Windows.Forms.ListBox();
            this.listView0 = new System.Windows.Forms.ListView();
            this.btnOk = new System.Windows.Forms.Button();
            this.tbSearch = new System.Windows.Forms.TextBox();
            this.сheckBox1 = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // // 
            // // listBox0
            // // 
            // this.listBox0.ColumnWidth = 440;
            // this.listBox0.ItemHeight = 15;
            // this.listBox0.Location = new System.Drawing.Point(16, 48);
            // this.listBox0.MultiColumn = false;
            // this.listBox0.Name = "listBox0";
            // this.listBox0.Size = new System.Drawing.Size(684,259);
            // this.listBox0.TabIndex = 0;
            // this.listBox0.DoubleClick += new System.EventHandler(this.listBox0_DoubleClick);
            // 
            // listView0
            // 
            // this.listView0.ColumnWidth = 440;
            // this.listView0.ItemHeight = 15;
            this.listView0.Location = new System.Drawing.Point(16, 48);
            // this.listView0.MultiColumn = false;
            this.listView0.Alignment = ListViewAlignment.Top;
            this.listView0.View = View.Details;
            this.listView0.Columns.Add("Формулы",660);
            this.listView0.MultiSelect = false;
            this.listView0.HeaderStyle = ColumnHeaderStyle.None;
            this.listView0.Name = "listView0";
            this.listView0.Size = new System.Drawing.Size(684,259);
            this.listView0.TabIndex = 0;
            this.listView0.DoubleClick += new System.EventHandler(this.listBox0_DoubleClick);
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(626,320);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 1;
            this.btnOk.Text = "Ok";
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // tbSearch
            // 
            this.tbSearch.Location = new System.Drawing.Point(16, 12);
            this.tbSearch.Name = "tbSearch";
            this.tbSearch.PlaceholderText = "Поиск...";
            this.tbSearch.Size = new System.Drawing.Size(596,23);
            this.tbSearch.TabIndex = 2;
            this.tbSearch.TextChanged += new System.EventHandler(this.tbSearch_TextChanged);
            // 
            // сheckBox1
            // 
            this.сheckBox1.AutoSize = true;
            this.сheckBox1.Location = new System.Drawing.Point(624,16);
            this.сheckBox1.Name = "сheckBox1";
            this.сheckBox1.Size = new System.Drawing.Size(69, 19);
            this.сheckBox1.TabIndex = 3;
            this.сheckBox1.Text = "Регистр";
            // 
            // AllFormulas
            // 
            this.ClientSize = new System.Drawing.Size(736,404);
            // this.Controls.Add(this.listBox0);
            this.Controls.Add(this.listView0);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.tbSearch);
            this.Controls.Add(this.сheckBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Name = "AllFormulas";
            this.Text = "Все формулы";
            this.Shown += new System.EventHandler(this.AllFormulas_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        } 

        #endregion 

        // protected ListBox listBox0;
        protected ListView listView0;
        protected Button btnOk;
        protected TextBox tbSearch;
        public CheckBox сheckBox1;
    }
}
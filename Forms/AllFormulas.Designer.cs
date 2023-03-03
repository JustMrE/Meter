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
            this.listBox0 = new ListBox();
            this.btnOk = new Button();
            this.tbSearch = new TextBox();
            this.сheckBox1 = new CheckBox();
            this.SuspendLayout();
            //
            // listBox0
            //
            this.listBox0.ItemHeight = 15;
            this.listBox0.MultiColumn =  true;
            this.listBox0.Text =  "listBox0";
            this.listBox0.Location = new System.Drawing.Point(16,48);
            this.listBox0.Size = new System.Drawing.Size(440,259);
            this.listBox0.ColumnWidth = listBox0.Size.Width;
            this.listBox0.DoubleClick += new System.EventHandler(listBox0_DoubleClick);
            //
            // btnOk
            //
            this.btnOk.Text =  "Ok";
            this.btnOk.Location = new System.Drawing.Point(382,320);
            this.btnOk.TabIndex = 1;
            this.btnOk.Click += new System.EventHandler(btnOk_Click);
            //
            // tbSearch
            //
            this.tbSearch.PlaceholderText = "Поиск...";
            this.tbSearch.Location = new System.Drawing.Point(16,12);
            this.tbSearch.Size = new System.Drawing.Size(356,23);
            this.tbSearch.TabIndex = 2;
            this.tbSearch.TextChanged += new System.EventHandler(tbSearch_TextChanged);
            //
            // сheckBox1
            //
            this.сheckBox1.AutoSize =  true;
            this.сheckBox1.Text =  "Регистр";
            this.сheckBox1.Location = new System.Drawing.Point(384,16);
            this.сheckBox1.Size = new System.Drawing.Size(69,19);
            this.сheckBox1.TabIndex = 3;
         //
         // form
         //
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Size = new System.Drawing.Size(492,404);
            this.Text =  "Все формулы";
            this.Controls.Add(this.listBox0);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.tbSearch);
            this.Controls.Add(this.сheckBox1);
            this.ResumeLayout(false);
        } 

        #endregion 

        protected ListBox listBox0;
        protected Button btnOk;
        protected TextBox tbSearch;
        public CheckBox сheckBox1;
    }
}
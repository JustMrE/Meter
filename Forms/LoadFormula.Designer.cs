namespace Meter
{
    partial class LoadFormula
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
            this.ListBox1 = new System.Windows.Forms.ListBox();
            this.tbSearch = new System.Windows.Forms.TextBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.SuspendLayout();
            //
            // ListBox1
            //
            this.ListBox1.ItemHeight = 15;
            this.ListBox1.Text =  "ListBox";
            this.ListBox1.Location = new System.Drawing.Point(12,36);
            this.ListBox1.Size = new System.Drawing.Size(372,184);
            this.ListBox1.TabIndex = 1;
            this.ListBox1.Click += new System.EventHandler(ListBox1_Click);
            //
            // tbSearch
            //
            this.tbSearch.Location = new System.Drawing.Point(12,8);
            this.tbSearch.Size = new System.Drawing.Size(368,23);
            this.tbSearch.TabIndex = 2;
            this.tbSearch.TextChanged += new System.EventHandler(this.tbSearch_TextChanged);
            //
            // btnOK
            //
            this.btnOK.Text =  "Ок";
            this.btnOK.Location = new System.Drawing.Point(224,260);
            this.btnOK.TabIndex = 3;
            this.btnOK.Click += new System.EventHandler(btnOK_Click);
            //
            // btnCancel
            //
            this.btnCancel.Text =  "Отмена";
            this.btnCancel.Location = new System.Drawing.Point(304,260);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Click += new System.EventHandler(btnCancel_Click);
            //
            // btnDelete
            //
            this.btnDelete.Text =  "Удалить";
            this.btnDelete.Enabled =  false;
            this.btnDelete.Location = new System.Drawing.Point(16,224);
            this.btnDelete.TabIndex = 5;
            this.btnDelete.Click += new System.EventHandler(btnDelete_Click);
         //
         // form
         //
            this.Size = new System.Drawing.Size(408,332);
            this.Text =  "Сохранение формулы";
            this.Controls.Add(this.ListBox1);
            this.Controls.Add(this.tbSearch);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnDelete);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ResumeLayout(false);
        } 

        #endregion 

        private System.Windows.Forms.ListBox ListBox1;
        private System.Windows.Forms.TextBox tbSearch;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnDelete;
    }
}


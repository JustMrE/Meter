namespace Meter
{
    partial class SaveFormula
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
            this.Label0 = new System.Windows.Forms.Label();
            this.ListBox1 = new System.Windows.Forms.ListBox();
            this.tbName = new System.Windows.Forms.TextBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.SuspendLayout();
            //
            // Label0
            //
            this.Label0.AutoSize =  true;
            this.Label0.Text =  "Введите название";
            this.Label0.Location = new System.Drawing.Point(12,212);
            this.Label0.Size = new System.Drawing.Size(103,15);
            //
            // ListBox1
            //
            this.ListBox1.ItemHeight = 15;
            this.ListBox1.Text =  "ListBox1";
            this.ListBox1.Location = new System.Drawing.Point(12,8);
            this.ListBox1.Size = new System.Drawing.Size(372,139);
            this.ListBox1.TabIndex = 1;
            this.ListBox1.Click += new System.EventHandler(ListBox1_Click);
            //
            // tbName
            //
            this.tbName.Text =  "tbName";
            this.tbName.Location = new System.Drawing.Point(12,236);
            this.tbName.Size = new System.Drawing.Size(368,23);
            this.tbName.TabIndex = 2;
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
            this.btnDelete.Location = new System.Drawing.Point(12,152);
            this.btnDelete.TabIndex = 5;
            this.btnDelete.Click += new System.EventHandler(btnDelete_Click);
         //
         // form
         //
            this.Size = new System.Drawing.Size(408,332);
            this.Text =  "Сохранение формулы";
            this.Controls.Add(this.Label0);
            this.Controls.Add(this.ListBox1);
            this.Controls.Add(this.tbName);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnDelete);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ResumeLayout(false);
        } 

        #endregion 

        private System.Windows.Forms.Label Label0;
        private System.Windows.Forms.ListBox ListBox1;
        private System.Windows.Forms.TextBox tbName;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnDelete;
    }
}

// private void ListBox1_Click(System.Object? sender, System.EventArgs e)
// {
// 
// }

// private void btnOK_Click(System.Object? sender, System.EventArgs e)
// {
// 
// }

// private void btnCancel_Click(System.Object? sender, System.EventArgs e)
// {
// 
// }

// private void btnDelete_Click(System.Object? sender, System.EventArgs e)
// {
// 
// }


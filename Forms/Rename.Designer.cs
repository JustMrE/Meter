namespace Meter
{
    partial class Rename
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
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.tbNewName = new System.Windows.Forms.TextBox();
            this.Label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            //
            // btnOk
            //
            this.btnOk.Text =  "Ok";
            this.btnOk.Location = new System.Drawing.Point(92,84);
            this.btnOk.Click += new System.EventHandler(btnOk_Click);
            //
            // btnCancel
            //
            this.btnCancel.Text =  "Cancel";
            this.btnCancel.Location = new System.Drawing.Point(224,84);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Click += new System.EventHandler(btnCancel_Click);
            //
            // tbNewName
            //
            this.tbNewName.Text =  "tbNewName";
            this.tbNewName.Location = new System.Drawing.Point(108,20);
            this.tbNewName.Size = new System.Drawing.Size(264,23);
            this.tbNewName.TabIndex = 2;
            //
            // Label4
            //
            this.Label4.AutoSize =  true;
            this.Label4.Text =  "Название";
            this.Label4.Location = new System.Drawing.Point(28,24);
            this.Label4.Size = new System.Drawing.Size(59,15);
            this.Label4.TabIndex = 4;
         //
         // form
         //
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Size = new System.Drawing.Size(400,168);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text =  "Переименование";
            this.Shown += new System.EventHandler(Rename_Shown);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.tbNewName);
            this.Controls.Add(this.Label4);
            this.ResumeLayout(false);
        } 

        #endregion 

        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox tbNewName;
        private System.Windows.Forms.Label Label4;
    }
}

// private void btnOk_Click(System.Object? sender, System.EventArgs e)
// {
// 
// }

// private void btnCancel_Click(System.Object? sender, System.EventArgs e)
// {
// 
// }

// private void Rename_Shown(System.Object? sender, System.EventArgs e)
// {
//
// }


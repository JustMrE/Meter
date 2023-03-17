namespace Meter.Forms
{
    public partial class DBClearing
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
            this.Label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            //
            // Label0
            //
            this.Label0.AutoSize =  true;
            this.Label0.Text =  "Очистка данных: ";
            this.Label0.BackColor = System.Drawing.SystemColors.Info;
            this.Label0.Font = new System.Drawing.Font("Times New Roman", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.Label0.Location = new System.Drawing.Point(20,48);
            this.Label0.Size = new System.Drawing.Size(134,17);
            //
            // Label1
            //
            this.Label1.AutoEllipsis =  true;
            this.Label1.Text =  "Label1";
            this.Label1.BackColor = System.Drawing.SystemColors.Info;
            this.Label1.Font = new System.Drawing.Font("Times New Roman", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.Label1.Location = new System.Drawing.Point(164,48);
            this.Label1.MaximumSize = new System.Drawing.Size(260,52);
            this.Label1.Size = new System.Drawing.Size(260,52);
            this.Label1.TabIndex = 1;
         //
         // form
         //
            this.BackColor = System.Drawing.SystemColors.Info;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Size = new System.Drawing.Size(456,120);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text =  "Form1";
            this.Controls.Add(this.Label0);
            this.Controls.Add(this.Label1);
            this.ResumeLayout(false);
        } 

        #endregion 

        private System.Windows.Forms.Label Label0;
        private System.Windows.Forms.Label Label1;
    }
}


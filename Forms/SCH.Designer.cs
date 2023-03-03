namespace Meter.Forms
{
    partial class SCH
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lCoef = new System.Windows.Forms.Label();
            this.tbPrev = new System.Windows.Forms.TextBox();
            this.tbNext = new System.Windows.Forms.TextBox();
            this.lSum = new System.Windows.Forms.Label();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(106, 19);
            this.label1.TabIndex = 0;
            this.label1.Text = "Коэффициент";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label2.Location = new System.Drawing.Point(12, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(105, 19);
            this.label2.TabIndex = 1;
            this.label2.Text = "Предыдущий";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label3.Location = new System.Drawing.Point(12, 67);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 19);
            this.label3.TabIndex = 2;
            this.label3.Text = "Текущий";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label4.Location = new System.Drawing.Point(12, 95);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 19);
            this.label4.TabIndex = 3;
            this.label4.Text = "Расход";
            // 
            // lCoef
            // 
            this.lCoef.AutoSize = true;
            this.lCoef.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.lCoef.Location = new System.Drawing.Point(135, 9);
            this.lCoef.Name = "lCoef";
            this.lCoef.Size = new System.Drawing.Size(33, 19);
            this.lCoef.TabIndex = 4;
            this.lCoef.Text = "999";
            // 
            // tbPrev
            // 
            this.tbPrev.Location = new System.Drawing.Point(135, 37);
            this.tbPrev.Name = "tbPrev";
            this.tbPrev.Size = new System.Drawing.Size(100, 23);
            this.tbPrev.TabIndex = 5;
            this.tbPrev.TextChanged += new System.EventHandler(this.tbPrev_TextChanged);
            this.tbPrev.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.tbPrev_KeyPress);
            // 
            // tbNext
            // 
            this.tbNext.Location = new System.Drawing.Point(135, 66);
            this.tbNext.Name = "tbNext";
            this.tbNext.Size = new System.Drawing.Size(100, 23);
            this.tbNext.TabIndex = 6;
            this.tbNext.TextChanged += new System.EventHandler(this.tbNext_TextChanged);
            this.tbNext.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.tbNext_KeyPress);
            // 
            // lSum
            // 
            this.lSum.AutoSize = true;
            this.lSum.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.lSum.Location = new System.Drawing.Point(135, 95);
            this.lSum.Name = "lSum";
            this.lSum.Size = new System.Drawing.Size(33, 19);
            this.lSum.TabIndex = 7;
            this.lSum.Text = "999";
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(42, 145);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 8;
            this.btnOk.Text = "Ok";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(135, 145);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 9;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // SCH
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(259, 190);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lCoef);
            this.Controls.Add(this.tbPrev);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lSum);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tbNext);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Name = "SCH";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SCH";
            this.Shown += new System.EventHandler(this.SCH_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label lCoef;
        private TextBox tbPrev;
        private TextBox tbNext;
        private Label lSum;
        private Button btnOk;
        private Button btnCancel;
    }
}
namespace Meter.Forms
{
    partial class Setup
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Setup));
            gbDB = new GroupBox();
            tbDB = new TextBox();
            btnSetDB = new Button();
            label1 = new Label();
            btnOk = new Button();
            btnCancel = new Button();
            flowLayoutPanel1 = new FlowLayoutPanel();
            gbMeter = new GroupBox();
            tbMeter = new TextBox();
            btnSetMeter = new Button();
            gbLogPath = new GroupBox();
            tbLogPath = new TextBox();
            btnSetLogPath = new Button();
            gbOkCancel = new GroupBox();
            gbDB.SuspendLayout();
            flowLayoutPanel1.SuspendLayout();
            gbMeter.SuspendLayout();
            gbLogPath.SuspendLayout();
            gbOkCancel.SuspendLayout();
            SuspendLayout();
            // 
            // gbDB
            // 
            gbDB.Controls.Add(tbDB);
            gbDB.Controls.Add(btnSetDB);
            gbDB.Location = new Point(13, 73);
            gbDB.Name = "gbDB";
            gbDB.Size = new Size(696, 56);
            gbDB.TabIndex = 3;
            gbDB.TabStop = false;
            gbDB.Text = "Путь к БД";
            // 
            // tbDB
            // 
            tbDB.Enabled = false;
            tbDB.Location = new Point(6, 22);
            tbDB.Name = "tbDB";
            tbDB.Size = new Size(647, 23);
            tbDB.TabIndex = 1;
            // 
            // btnSetDB
            // 
            btnSetDB.Location = new Point(659, 22);
            btnSetDB.Name = "btnSetDB";
            btnSetDB.Size = new Size(27, 23);
            btnSetDB.TabIndex = 0;
            btnSetDB.Text = "...";
            btnSetDB.UseVisualStyleBackColor = true;
            btnSetDB.Click += btnSetDB_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            label1.ForeColor = Color.Red;
            label1.Location = new Point(13, 10);
            label1.Name = "label1";
            label1.Size = new Size(629, 60);
            label1.TabIndex = 4;
            label1.Text = resources.GetString("label1.Text");
            // 
            // btnOk
            // 
            btnOk.Location = new Point(530, 22);
            btnOk.Name = "btnOk";
            btnOk.Size = new Size(75, 23);
            btnOk.TabIndex = 6;
            btnOk.Text = "Ok";
            btnOk.UseVisualStyleBackColor = true;
            btnOk.Click += btnOk_Click;
            // 
            // btnCancel
            // 
            btnCancel.Location = new Point(611, 22);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(75, 23);
            btnCancel.TabIndex = 5;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = true;
            btnCancel.Click += btnCancel_Click;
            // 
            // flowLayoutPanel1
            // 
            flowLayoutPanel1.AutoSize = true;
            flowLayoutPanel1.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            flowLayoutPanel1.Controls.Add(label1);
            flowLayoutPanel1.Controls.Add(gbDB);
            flowLayoutPanel1.Controls.Add(gbMeter);
            flowLayoutPanel1.Controls.Add(gbLogPath);
            flowLayoutPanel1.Controls.Add(gbOkCancel);
            flowLayoutPanel1.FlowDirection = FlowDirection.TopDown;
            flowLayoutPanel1.Location = new Point(18, 12);
            flowLayoutPanel1.Name = "flowLayoutPanel1";
            flowLayoutPanel1.Padding = new Padding(10);
            flowLayoutPanel1.Size = new Size(722, 328);
            flowLayoutPanel1.TabIndex = 7;
            // 
            // gbMeter
            // 
            gbMeter.Controls.Add(tbMeter);
            gbMeter.Controls.Add(btnSetMeter);
            gbMeter.Enabled = false;
            gbMeter.Location = new Point(13, 135);
            gbMeter.Name = "gbMeter";
            gbMeter.Size = new Size(696, 56);
            gbMeter.TabIndex = 9;
            gbMeter.TabStop = false;
            gbMeter.Text = "Файл Excel";
            gbMeter.Visible = false;
            // 
            // tbMeter
            // 
            tbMeter.Enabled = false;
            tbMeter.Location = new Point(6, 22);
            tbMeter.Name = "tbMeter";
            tbMeter.Size = new Size(647, 23);
            tbMeter.TabIndex = 1;
            // 
            // btnSetMeter
            // 
            btnSetMeter.Location = new Point(659, 22);
            btnSetMeter.Name = "btnSetMeter";
            btnSetMeter.Size = new Size(27, 23);
            btnSetMeter.TabIndex = 0;
            btnSetMeter.Text = "...";
            btnSetMeter.UseVisualStyleBackColor = true;
            // 
            // gbLogPath
            // 
            gbLogPath.Controls.Add(tbLogPath);
            gbLogPath.Controls.Add(btnSetLogPath);
            gbLogPath.Enabled = false;
            gbLogPath.Location = new Point(13, 197);
            gbLogPath.Name = "gbLogPath";
            gbLogPath.Size = new Size(696, 56);
            gbLogPath.TabIndex = 10;
            gbLogPath.TabStop = false;
            gbLogPath.Text = "Путь к Log файлам";
            gbLogPath.Visible = false;
            // 
            // tbLogPath
            // 
            tbLogPath.Enabled = false;
            tbLogPath.Location = new Point(6, 22);
            tbLogPath.Name = "tbLogPath";
            tbLogPath.Size = new Size(647, 23);
            tbLogPath.TabIndex = 1;
            // 
            // btnSetLogPath
            // 
            btnSetLogPath.Location = new Point(659, 22);
            btnSetLogPath.Name = "btnSetLogPath";
            btnSetLogPath.Size = new Size(27, 23);
            btnSetLogPath.TabIndex = 0;
            btnSetLogPath.Text = "...";
            btnSetLogPath.UseVisualStyleBackColor = true;
            // 
            // gbOkCancel
            // 
            gbOkCancel.Controls.Add(btnCancel);
            gbOkCancel.Controls.Add(btnOk);
            gbOkCancel.Location = new Point(13, 259);
            gbOkCancel.Name = "gbOkCancel";
            gbOkCancel.Size = new Size(696, 56);
            gbOkCancel.TabIndex = 8;
            gbOkCancel.TabStop = false;
            // 
            // Setup
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            ClientSize = new Size(869, 436);
            Controls.Add(flowLayoutPanel1);
            FormBorderStyle = FormBorderStyle.Fixed3D;
            Name = "Setup";
            ShowIcon = false;
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Setup";
            gbDB.ResumeLayout(false);
            gbDB.PerformLayout();
            flowLayoutPanel1.ResumeLayout(false);
            flowLayoutPanel1.PerformLayout();
            gbMeter.ResumeLayout(false);
            gbMeter.PerformLayout();
            gbLogPath.ResumeLayout(false);
            gbLogPath.PerformLayout();
            gbOkCancel.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private GroupBox gbDB;
        private TextBox tbDB;
        private Button btnSetDB;
        private Label label1;
        private Button btnOk;
        private Button btnCancel;
        private FlowLayoutPanel flowLayoutPanel1;
        private GroupBox gbOkCancel;
        private GroupBox gbMeter;
        private TextBox tbMeter;
        private Button btnSetMeter;
        private GroupBox gbLogPath;
        private TextBox tbLogPath;
        private Button btnSetLogPath;
    }
}
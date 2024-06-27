namespace Meter.Forms
{
    partial class Settings
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
            tabControl1 = new TabControl();
            tabPage1 = new TabPage();
            cbCloseSaveResponce = new CheckBox();
            cbSaveOnClose = new CheckBox();
            gbLogPath = new GroupBox();
            btnResetLogPath = new Button();
            tbLogPath = new TextBox();
            btnSetLogPath = new Button();
            gbMeter = new GroupBox();
            btnResetMeter = new Button();
            tbMeter = new TextBox();
            btnSetMeter = new Button();
            gbDB = new GroupBox();
            btnResetDB = new Button();
            tbDB = new TextBox();
            btnSetDB = new Button();
            tabPage2 = new TabPage();
            gbEmcosConnection = new GroupBox();
            tbEmcosHost = new TextBox();
            lEmcosUrl = new Label();
            lEmcosHost = new Label();
            tbEmcosUrl = new TextBox();
            gbAuth = new GroupBox();
            tbEmcosPass = new TextBox();
            lEmcosLogin = new Label();
            lEmcosPass = new Label();
            tbEmcosLogin = new TextBox();
            btnCancel = new Button();
            btnOk = new Button();
            tabControl1.SuspendLayout();
            tabPage1.SuspendLayout();
            gbLogPath.SuspendLayout();
            gbMeter.SuspendLayout();
            gbDB.SuspendLayout();
            tabPage2.SuspendLayout();
            gbEmcosConnection.SuspendLayout();
            gbAuth.SuspendLayout();
            SuspendLayout();
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Location = new Point(12, 12);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(776, 397);
            tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            tabPage1.Controls.Add(cbCloseSaveResponce);
            tabPage1.Controls.Add(cbSaveOnClose);
            tabPage1.Controls.Add(gbLogPath);
            tabPage1.Controls.Add(gbMeter);
            tabPage1.Controls.Add(gbDB);
            tabPage1.Location = new Point(4, 24);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(768, 369);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "Общие";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // cbCloseSaveResponce
            // 
            cbCloseSaveResponce.AutoSize = true;
            cbCloseSaveResponce.Location = new Point(214, 6);
            cbCloseSaveResponce.Name = "cbCloseSaveResponce";
            cbCloseSaveResponce.Size = new Size(242, 19);
            cbCloseSaveResponce.TabIndex = 6;
            cbCloseSaveResponce.Text = "Спрашивать сохранение при закрытии";
            cbCloseSaveResponce.UseVisualStyleBackColor = true;
            cbCloseSaveResponce.CheckedChanged += cbCloseSaveResponce_CheckedChanged;
            // 
            // cbSaveOnClose
            // 
            cbSaveOnClose.AutoSize = true;
            cbSaveOnClose.Location = new Point(12, 6);
            cbSaveOnClose.Name = "cbSaveOnClose";
            cbSaveOnClose.Size = new Size(196, 19);
            cbSaveOnClose.TabIndex = 5;
            cbSaveOnClose.Text = "Автосохранение при закрытии";
            cbSaveOnClose.UseVisualStyleBackColor = true;
            cbSaveOnClose.CheckedChanged += cbSaveOnClose_CheckedChanged;
            // 
            // gbLogPath
            // 
            gbLogPath.Controls.Add(btnResetLogPath);
            gbLogPath.Controls.Add(tbLogPath);
            gbLogPath.Controls.Add(btnSetLogPath);
            gbLogPath.Location = new Point(6, 155);
            gbLogPath.Name = "gbLogPath";
            gbLogPath.Size = new Size(756, 56);
            gbLogPath.TabIndex = 4;
            gbLogPath.TabStop = false;
            gbLogPath.Text = "Путь к Log файлам";
            // 
            // btnResetLogPath
            // 
            btnResetLogPath.Location = new Point(692, 22);
            btnResetLogPath.Name = "btnResetLogPath";
            btnResetLogPath.Size = new Size(58, 23);
            btnResetLogPath.TabIndex = 3;
            btnResetLogPath.Text = "сброс";
            btnResetLogPath.UseVisualStyleBackColor = true;
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
            btnSetLogPath.Click += btnSetLogPath_Click;
            // 
            // gbMeter
            // 
            gbMeter.Controls.Add(btnResetMeter);
            gbMeter.Controls.Add(tbMeter);
            gbMeter.Controls.Add(btnSetMeter);
            gbMeter.Enabled = false;
            gbMeter.Location = new Point(6, 93);
            gbMeter.Name = "gbMeter";
            gbMeter.Size = new Size(756, 56);
            gbMeter.TabIndex = 3;
            gbMeter.TabStop = false;
            gbMeter.Text = "Файл Excel";
            // 
            // btnResetMeter
            // 
            btnResetMeter.Location = new Point(692, 22);
            btnResetMeter.Name = "btnResetMeter";
            btnResetMeter.Size = new Size(58, 23);
            btnResetMeter.TabIndex = 3;
            btnResetMeter.Text = "сброс";
            btnResetMeter.UseVisualStyleBackColor = true;
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
            // gbDB
            // 
            gbDB.Controls.Add(btnResetDB);
            gbDB.Controls.Add(tbDB);
            gbDB.Controls.Add(btnSetDB);
            gbDB.Enabled = false;
            gbDB.Location = new Point(6, 31);
            gbDB.Name = "gbDB";
            gbDB.Size = new Size(756, 56);
            gbDB.TabIndex = 2;
            gbDB.TabStop = false;
            gbDB.Text = "Путь к БД";
            // 
            // btnResetDB
            // 
            btnResetDB.Location = new Point(692, 22);
            btnResetDB.Name = "btnResetDB";
            btnResetDB.Size = new Size(58, 23);
            btnResetDB.TabIndex = 2;
            btnResetDB.Text = "сброс";
            btnResetDB.UseVisualStyleBackColor = true;
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
            btnSetDB.BackgroundImageLayout = ImageLayout.None;
            btnSetDB.Location = new Point(659, 22);
            btnSetDB.Name = "btnSetDB";
            btnSetDB.Size = new Size(27, 23);
            btnSetDB.TabIndex = 0;
            btnSetDB.Text = "...";
            btnSetDB.UseVisualStyleBackColor = true;
            btnSetDB.Click += btnSetDB_Click;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(gbEmcosConnection);
            tabPage2.Controls.Add(gbAuth);
            tabPage2.Location = new Point(4, 24);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(768, 369);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "EMCOS";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // gbEmcosConnection
            // 
            gbEmcosConnection.Controls.Add(tbEmcosHost);
            gbEmcosConnection.Controls.Add(lEmcosUrl);
            gbEmcosConnection.Controls.Add(lEmcosHost);
            gbEmcosConnection.Controls.Add(tbEmcosUrl);
            gbEmcosConnection.Location = new Point(310, 6);
            gbEmcosConnection.Name = "gbEmcosConnection";
            gbEmcosConnection.Size = new Size(452, 125);
            gbEmcosConnection.TabIndex = 1;
            gbEmcosConnection.TabStop = false;
            gbEmcosConnection.Text = "Соединение";
            // 
            // tbEmcosHost
            // 
            tbEmcosHost.Location = new Point(6, 93);
            tbEmcosHost.Name = "tbEmcosHost";
            tbEmcosHost.Size = new Size(440, 23);
            tbEmcosHost.TabIndex = 8;
            tbEmcosHost.TextChanged += tbEmcosHost_TextChanged;
            // 
            // lEmcosUrl
            // 
            lEmcosUrl.AutoSize = true;
            lEmcosUrl.Location = new Point(6, 19);
            lEmcosUrl.Name = "lEmcosUrl";
            lEmcosUrl.Size = new Size(28, 15);
            lEmcosUrl.TabIndex = 5;
            lEmcosUrl.Text = "URL";
            // 
            // lEmcosHost
            // 
            lEmcosHost.AutoSize = true;
            lEmcosHost.Location = new Point(6, 75);
            lEmcosHost.Name = "lEmcosHost";
            lEmcosHost.Size = new Size(32, 15);
            lEmcosHost.TabIndex = 6;
            lEmcosHost.Text = "Host";
            // 
            // tbEmcosUrl
            // 
            tbEmcosUrl.Location = new Point(6, 37);
            tbEmcosUrl.Name = "tbEmcosUrl";
            tbEmcosUrl.Size = new Size(440, 23);
            tbEmcosUrl.TabIndex = 7;
            tbEmcosUrl.TextChanged += tbEmcosUrl_TextChanged;
            // 
            // gbAuth
            // 
            gbAuth.Controls.Add(tbEmcosPass);
            gbAuth.Controls.Add(lEmcosLogin);
            gbAuth.Controls.Add(lEmcosPass);
            gbAuth.Controls.Add(tbEmcosLogin);
            gbAuth.Location = new Point(6, 6);
            gbAuth.Name = "gbAuth";
            gbAuth.Size = new Size(298, 125);
            gbAuth.TabIndex = 0;
            gbAuth.TabStop = false;
            gbAuth.Text = "Авторизация";
            // 
            // tbEmcosPass
            // 
            tbEmcosPass.Location = new Point(6, 93);
            tbEmcosPass.Name = "tbEmcosPass";
            tbEmcosPass.Size = new Size(284, 23);
            tbEmcosPass.TabIndex = 4;
            tbEmcosPass.TextChanged += tbEmcosPass_TextChanged;
            // 
            // lEmcosLogin
            // 
            lEmcosLogin.AutoSize = true;
            lEmcosLogin.Location = new Point(6, 19);
            lEmcosLogin.Name = "lEmcosLogin";
            lEmcosLogin.Size = new Size(41, 15);
            lEmcosLogin.TabIndex = 1;
            lEmcosLogin.Text = "Логин";
            // 
            // lEmcosPass
            // 
            lEmcosPass.AutoSize = true;
            lEmcosPass.Location = new Point(6, 75);
            lEmcosPass.Name = "lEmcosPass";
            lEmcosPass.Size = new Size(49, 15);
            lEmcosPass.TabIndex = 2;
            lEmcosPass.Text = "Пароль";
            // 
            // tbEmcosLogin
            // 
            tbEmcosLogin.Location = new Point(6, 37);
            tbEmcosLogin.Name = "tbEmcosLogin";
            tbEmcosLogin.Size = new Size(284, 23);
            tbEmcosLogin.TabIndex = 3;
            tbEmcosLogin.TextChanged += tbEmcosLogin_TextChanged;
            // 
            // btnCancel
            // 
            btnCancel.Location = new Point(713, 415);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(75, 23);
            btnCancel.TabIndex = 1;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = true;
            btnCancel.Click += btnCancel_Click;
            // 
            // btnOk
            // 
            btnOk.Location = new Point(632, 415);
            btnOk.Name = "btnOk";
            btnOk.Size = new Size(75, 23);
            btnOk.TabIndex = 2;
            btnOk.Text = "Ok";
            btnOk.UseVisualStyleBackColor = true;
            btnOk.Click += btnOk_Click;
            // 
            // Settings
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(btnOk);
            Controls.Add(btnCancel);
            Controls.Add(tabControl1);
            FormBorderStyle = FormBorderStyle.Fixed3D;
            Name = "Settings";
            ShowIcon = false;
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Settings";
            Shown += Settings_Shown;
            tabControl1.ResumeLayout(false);
            tabPage1.ResumeLayout(false);
            tabPage1.PerformLayout();
            gbLogPath.ResumeLayout(false);
            gbLogPath.PerformLayout();
            gbMeter.ResumeLayout(false);
            gbMeter.PerformLayout();
            gbDB.ResumeLayout(false);
            gbDB.PerformLayout();
            tabPage2.ResumeLayout(false);
            gbEmcosConnection.ResumeLayout(false);
            gbEmcosConnection.PerformLayout();
            gbAuth.ResumeLayout(false);
            gbAuth.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private TabControl tabControl1;
        private TabPage tabPage1;
        private TextBox tbDB;
        private Button btnSetDB;
        private TabPage tabPage2;
        private GroupBox gbDB;
        private GroupBox gbMeter;
        private TextBox tbMeter;
        private Button btnSetMeter;
        private GroupBox gbLogPath;
        private Button btnResetLogPath;
        private TextBox tbLogPath;
        private Button btnSetLogPath;
        private Button btnResetMeter;
        private Button btnResetDB;
        private Button btnCancel;
        private Button btnOk;
        private CheckBox cbCloseSaveResponce;
        private CheckBox cbSaveOnClose;
        private TextBox tbEmcosPass;
        private Label lEmcosPass;
        private GroupBox gbAuth;
        private Label lEmcosLogin;
        private TextBox tbEmcosLogin;
        private GroupBox gbEmcosConnection;
        private TextBox tbEmcosHost;
        private Label lEmcosUrl;
        private Label lEmcosHost;
        private TextBox tbEmcosUrl;
    }
}
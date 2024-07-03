using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Meter.Forms
{
    public partial class Settings : Form
    {
        string DBDir { get; set; }
        string MeterFile { get; set; }
        string LogFilePath { get; set; }
        bool CloseAutoSave { get; set; }
        bool CloseSaveResponce { get; set; }
        string EMCOSLogin { get; set; }
        string EMCOSPassword { get; set; }
        string EmcosUrl { get; set; }
        string EmcosHost { get; set; }
        string OldMeter { get; set; }
        string OtherMetersPath { get; set; }

        public Settings()
        {
            InitializeComponent();
        }

        private void btnSetDB_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                string dbFolder = folderDialog.SelectedPath;
                string meterFile = dbFolder + @"\current\meter.xlsx";
                string logFilePath = dbFolder + @"\current\";
                if (!File.Exists(meterFile))
                {
                    MessageBox.Show("Не найден Excel файл в папке БД!\nБудет создан новый файл счетчиков!");
                    //this.tbDB.Text = string.Empty;
                    this.tbMeter.Text = string.Empty;
                    this.tbLogPath.Text = string.Empty;
                }
                else
                {
                    this.tbDB.Text = dbFolder;
                    this.tbMeter.Text = meterFile;
                    this.tbLogPath.Text = logFilePath;
                    MeterSettings.Instance.DBDir = dbFolder;
                    MeterSettings.Instance.MeterFile = meterFile;
                    MeterSettings.Instance.LogFile = logFilePath + "log.log";
                    MeterSettings.Instance.ErrLogFile = logFilePath + "errlog.log";
                }
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            //MeterSettings.Instance.DBDir = DBDir;
            //MeterSettings.Instance.MeterFile = MeterFile;
            MeterSettings.Instance.LogFile = LogFilePath + @"\log.log";
            MeterSettings.Instance.ErrLogFile = LogFilePath + @"\errlog.log";
            MeterSettings.Instance.CloseAutoSave = CloseAutoSave;
            MeterSettings.Instance.CloseSaveResponce = CloseSaveResponce;
            MeterSettings.Instance.EmcosLogin = EMCOSLogin;
            MeterSettings.Instance.EmcosPassword = EMCOSPassword;
            MeterSettings.Instance.EmcosUrl = EmcosUrl;
            MeterSettings.Instance.EmcosHost = EmcosHost;
            MeterSettings.Instance.OldMeterFile = OldMeter;
            MeterSettings.Instance.OtherMetersPath = OtherMetersPath;
            MeterSettings.Instance.EmcosUrl = EmcosUrl;
            MeterSettings.Instance.EmcosHost = EmcosHost;
            MeterSettings.Instance.Save();
            this.DialogResult = DialogResult.OK;
            Close();
        }

        private void Settings_Shown(object sender, EventArgs e)
        {
            this.tbDB.Text = DBDir = MeterSettings.Instance.DBDir;
            this.tbMeter.Text = MeterFile = MeterSettings.Instance.MeterFile;
            this.tbLogPath.Text = LogFilePath = Path.GetDirectoryName(MeterSettings.Instance.LogFile);
            this.cbSaveOnClose.Checked = CloseAutoSave = MeterSettings.Instance.CloseAutoSave;
            this.cbCloseSaveResponce.Checked = CloseSaveResponce = MeterSettings.Instance.CloseSaveResponce;

            this.tbEmcosLogin.Text = EMCOSLogin = MeterSettings.Instance.EmcosLogin;
            this.tbEmcosPass.Text = EMCOSPassword = MeterSettings.Instance.EmcosPassword;
            this.tbEmcosUrl.Text = EmcosUrl = MeterSettings.Instance.EmcosUrl;
            this.tbEmcosHost.Text = EmcosHost = MeterSettings.Instance.EmcosHost;
            this.tbOldMeter.Text = OldMeter = MeterSettings.Instance.OldMeterFile;
            this.tbMetersPath.Text = OtherMetersPath = MeterSettings.Instance.OtherMetersPath;
        }


        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            Close();
        }

        private void cbSaveOnClose_CheckedChanged(object sender, EventArgs e)
        {
            CloseAutoSave = cbSaveOnClose.Checked;
        }
        
        private void cbCloseSaveResponce_CheckedChanged(object sender, EventArgs e)
        {
            CloseSaveResponce = cbCloseSaveResponce.Checked;
        }

        private void tbEmcosLogin_TextChanged(object sender, EventArgs e)
        {
            EMCOSLogin = tbEmcosLogin.Text;
        }

        private void tbEmcosPass_TextChanged(object sender, EventArgs e)
        {
            EMCOSPassword = tbEmcosPass.Text;
        }

        private void btnSetLogPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                LogFilePath = folderDialog.SelectedPath;
            }
        }

        private void tbEmcosUrl_TextChanged(object sender, EventArgs e)
        {
            EmcosUrl = tbEmcosUrl.Text;
        }

        private void tbEmcosHost_TextChanged(object sender, EventArgs e)
        {
            EmcosHost = tbEmcosHost.Text;
        }
    
        private void btnSetMetersPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                OtherMetersPath = folderDialog.SelectedPath;
                this.tbMetersPath.Text = OtherMetersPath;
            }
        }

        private void btnSetOldMeter_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                OldMeter = openFileDialog.FileName;
                this.tbOldMeter.Text = OldMeter;
            }
        }
    
        
        private void btnEmcosUrl_Click(object sender, EventArgs e)
        {
            // this.tbEmcosUrl.Text = @"http://10.0.144.11:8080/ec3api/v1";//@"http://10.0.144.11:8080/emcos3/";
            string url = @"http://10.0.144.11:8080/ec3api/v1";
            Uri uri = new Uri(url);

            this.tbEmcosUrl.Text = uri.AbsoluteUri;
            this.tbEmcosHost.Text = uri.Host + ":" + uri.Port;
        }
    }
}

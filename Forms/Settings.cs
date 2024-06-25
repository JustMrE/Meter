using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Meter.Forms
{
    public partial class Settings : Form
    {
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
            MeterSettings.Instance.Save();
            this.DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}

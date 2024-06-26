using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Meter.Forms
{
    public partial class Setup : Form
    {
        bool newDB;
        string dbFolder;

        public Setup()
        {
            newDB = false;
            InitializeComponent();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (newDB)
            {
                CreateNewMetersDB();
            }
            else
            {

            }
            MeterSettings.Instance.Save();
            this.DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnSetDB_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                dbFolder = folderDialog.SelectedPath;
                string meterFile = dbFolder + @"\current\meter.xlsx";
                string logFilePath = dbFolder + @"\current";

                MeterSettings.Instance.DBDir = dbFolder;
                MeterSettings.Instance.MeterFile = meterFile;
                MeterSettings.Instance.LogFile = logFilePath + @"\log.log";
                MeterSettings.Instance.ErrLogFile = logFilePath + @"\errlog.log";
                
                if (!MeterSettings.Instance.CheckDBFiles())
                {
                    MessageBox.Show("Не найден Excel файл в папке БД, либо БД повреждена!\nБудет создан новый файл счетчиков\n(Скопируйте нужные не поврежденные данные перед продолжением!)","",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    this.tbDB.Text = dbFolder;
                    newDB = true;
                    this.gbMeter.Visible = false;
                    this.gbLogPath.Visible = false;
                }
                else
                {
                    MessageBox.Show("БД обнаружена!");
                    newDB = false;

                    this.gbMeter.Visible = true;
                    this.gbLogPath.Visible = true;

                    this.tbDB.Text = dbFolder;
                    this.tbMeter.Text = meterFile;
                    this.tbLogPath.Text = logFilePath;
                }
            }
        }

        private void CreateNewMetersDB()
        {
            // Очистка директории
            if (Directory.Exists(dbFolder))
            {
                Directory.Delete(dbFolder, true);
            }
            Directory.CreateDirectory(dbFolder);

            string arch = dbFolder + @"\arch";
            
            string current = dbFolder + @"\current";
            string formulas = current + @"\formulas";
            string references = current + @"\references";
            string meterFile = current + @"\meter.xlsx";
            string tiDictFile = current + @"\Словарь ТИ факт.xlsx";
            string colors = current + @"\colors.json";


            string saves = dbFolder + @"\saves";
            string savedFormulas = saves + @"\formulas";
            string tempArch = dbFolder + @"\temparch";

            string standartColors = dbFolder + @"\standartColors.json";

            Directory.CreateDirectory(arch);
            Directory.CreateDirectory(current);
            Directory.CreateDirectory(formulas);
            Directory.CreateDirectory(references);
            Directory.CreateDirectory(saves);
            Directory.CreateDirectory(savedFormulas);
            Directory.CreateDirectory(tempArch);

            File.WriteAllBytes(meterFile, Properties.Resources.meter);
            File.WriteAllBytes(tiDictFile, Properties.Resources.tiDict);
            File.WriteAllBytes(colors, Properties.Resources.standartColors);
            File.WriteAllBytes(standartColors, Properties.Resources.standartColors);

            MeterSettings.Instance.DBDir = dbFolder;
            MeterSettings.Instance.MeterFile = meterFile;
            MeterSettings.Instance.LogFile = current + @"\log.log";
            MeterSettings.Instance.ErrLogFile = current + @"\errlog.log";

            this.tbDB.Text = dbFolder;
            this.tbMeter.Text = meterFile;
            this.tbLogPath.Text = current;

            // this.gbMeter.Visible = true;
            // this.gbLogPath.Visible = true;
        }
    }
}

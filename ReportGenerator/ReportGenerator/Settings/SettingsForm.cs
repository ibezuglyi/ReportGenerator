using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace ReportGenerator.Settings
{
    public partial class SettingsForm : Form
    {
        public string SelectedConfigPath { get; set; }
        public string SelectedProfilePath { get; set; }
        public SettingsForm()
        {
            InitializeComponent();
            LoadData();
        }
        private void settingsForm_Load(object sender, EventArgs e)
        {

        }
        private void LoadData()
        {
            textBox1.Text = ReportConfiguration.Instance.ConfigurationFilePath;
            textBox2.Text = ReportConfiguration.Instance.ConfigurationProfileDirectory;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CloseForm();
        }

        private void CloseForm()
        {
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ReportConfiguration.Instance.ConfigurationFilePath = SelectedConfigPath;
            ReportConfiguration.Instance.ConfigurationProfileDirectory = SelectedProfilePath;
            CloseForm();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var dialogResult = openFileDialog1.ShowDialog();
            if (dialogResult == System.Windows.Forms.DialogResult.OK)
            {
                SelectedConfigPath = openFileDialog1.FileName;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var dialogResult = folderBrowserDialog1.ShowDialog();
            if (dialogResult == System.Windows.Forms.DialogResult.OK)
            {

            }
        }
    }
}

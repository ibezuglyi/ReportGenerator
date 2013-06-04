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

            ReportConfiguration.Instance.ConfigurationProfileDirectory = textBox2.Text;
            CloseForm();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            var dialogResult = folderBrowserDialog1.ShowDialog();
            if (dialogResult == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = folderBrowserDialog1.SelectedPath;
            }
        }
    }
}

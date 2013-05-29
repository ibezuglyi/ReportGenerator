using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportGenerator
{
    public class ReportConfiguration
    {
        private static ReportConfiguration instance;
        private ReportConfiguration()
        {
            ConfigurationFilePath = ReportGenerator.Properties.Settings.Default.ConfigurationPath;
            ConfigurationProfileDirectory = ReportGenerator.Properties.Settings.Default.ConfigurationProfileDirectory;
        }

        public static ReportConfiguration Instance
        {
            get
            {
                instance = instance ?? new ReportConfiguration();
                return instance;
            }
        }
        public string ConfigurationFilePath { get; set; }
        public string ConfigurationProfileDirectory { get; set; }

        private void LoadConfiguration()
        { }




    }
}

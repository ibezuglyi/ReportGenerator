using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace ReportGenerator.Converters
{
    public abstract class AssessmentConverter
    {
        private Configuration.Configuration currentConfiguration;

        public Assessment Convert(Assessment assessment, string configurationPath)
        {
            currentConfiguration = GetConfiguration(configurationPath);
            if (currentConfiguration == null)
                return null;

            return ConvertAssessment(assessment);
        }
        protected abstract Assessment ConvertAssessment(Assessment assessment);

        private Configuration.Configuration GetConfiguration(string configurationPath)
        {
            if (string.IsNullOrEmpty(configurationPath) || !File.Exists(configurationPath))
                return null;

            using (var stream = new StreamReader(configurationPath))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Configuration.Configuration));
                return serializer.Deserialize(stream) as Configuration.Configuration;
            }
        }
    }
}

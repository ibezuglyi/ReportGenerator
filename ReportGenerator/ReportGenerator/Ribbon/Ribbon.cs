using Microsoft.Office.Tools.Ribbon;
using ReportGenerator.Converters;
using System.IO;
using System.Xml.Serialization;

namespace ReportGenerator.Ribbon
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.SettingsForm settingsForm = new Settings.SettingsForm();
            settingsForm.Show();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            App.ReportGeneratorAddIn.ConvertAssessment(new NewAssessmentConverter());

            Serialize();

        }

        private void Serialize()
        {
            Configuration.Configuration c = new Configuration.Configuration();

            c.JavaDeveloper.AddMappings("Servlets", "Servlets");
            c.JavaDeveloper.AddMappings("Spring", "Spring");
            c.JavaDeveloper.KeyTechnologies.Add("Servlets");
            c.JavaDeveloper.KeyTechnologies.Add("OpenJPA");
            c.JavaDeveloper.KeyTechnologies.Add("Spring");
            c.Tester.AddMappings("installation", "installation1", "installation2");
            using (var stream = new StreamWriter(@"d:\stream.xml"))
            {
                XmlSerializer xml = new XmlSerializer(typeof(Configuration.Configuration));
                xml.Serialize(stream, c);
            }

          
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            App.ReportGeneratorAddIn.ConvertAssessment(new OldAssessmentConverter());

        }
    }
}

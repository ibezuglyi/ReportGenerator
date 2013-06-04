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
           
        }
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            App.ReportGeneratorAddIn.ConvertAssessment(new OldAssessmentConverter());

        }
    }
}

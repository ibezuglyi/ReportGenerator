using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using ReportGenerator.Converters;
using System.IO;
using System.Xml.Serialization;
using ReportGenerator.Profiles;

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

            EngineerProfile prof = new EngineerProfile();
            prof.OldProfile.TechnologyList.Add(new TechnologyItem(){ Technology= "JavaEE", Color=34567, isBold=true, isMerged=true });
            prof.OldProfile.TechnologyList.Add(new TechnologyItem() { Technology = "1) Multi-tier architecture and JavaEE technology", Color = 0, isBold = false, isMerged = false });
            prof.OldProfile.TechnologyList.Add(new TechnologyItem() { Technology = "2) JavaEE technologies", Color = 0, isBold = false, isMerged = false });
            prof.OldProfile.TechnologyList.Add(new TechnologyItem() { Technology = " JSP + HTML, ", Color = 0, isBold = false, isMerged = false });
            prof.NewProfile.TechnologyList.Add(new TechnologyItem() { Technology = "JavaSE", Color = 34567, isBold = true, isMerged = true });
            prof.NewProfile.TechnologyList.Add(new TechnologyItem() { Technology = "1) Java SE 1.4", Color = 0, isBold =true, isMerged = false });
            using (var stream = new StreamWriter(@"d:\profile.xml"))
            {
                XmlSerializer xml = new XmlSerializer(typeof(EngineerProfile));
                xml.Serialize(stream, prof);
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            App.ReportGeneratorAddIn.ConvertAssessment(new OldAssessmentConverter());

        }
    }
}

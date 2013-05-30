using ReportGenerator.Profiles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ReportGenerator.Console
{
    class Program
    {
        private static EngineerProfile BuildProfile()
        {
            EngineerProfile prof = new EngineerProfile();
            prof.Header.Scales.AddRange(

                new List<string>() { "Scale (0-4)", "0 - No experience.", "1 - Beginner (is able to perform simple task under supervision, usually less than 1 year of experience)",
                "2 - Autonomous (is able to perform task without supervision but may need support from expert, usually 1-2 year of experience)",
                "3 - Expert (is able to perform task without supervision, mentor/coach for others, usually 2-5 year of experience)",
                "4 - Guru"
                });
            var item = new TechnologyItem() { Technology = "JavaEE", Color = 34567, isBold = true, isMerged = true };
            item.MapToTechnologies.Add("JavaEE");
            prof.OldProfile.TechnologyList.Add(item);
            item = new TechnologyItem() { Technology = "1) Multi-tier architecture and JavaEE technology", Color = 0, isBold = false, isMerged = false };
            prof.OldProfile.TechnologyList.Add(item);
            prof.OldProfile.TechnologyList.Add(new TechnologyItem() { Technology = "2) JavaEE technologies", Color = 0, isBold = false, isMerged = false });
            item = new TechnologyItem() { Technology = " JSP + HTML, ", Color = 0, isBold = false, isMerged = false };
            item.MapToTechnologies.Add("JSP");
            prof.OldProfile.TechnologyList.Add(item);
            prof.NewProfile.TechnologyList.Add(new TechnologyItem() { Technology = "JavaSE", Color = 34567, isBold = true, isMerged = true });
            prof.NewProfile.TechnologyList.Add(new TechnologyItem() { Technology = "1) Java SE 1.4", Color = 0, isBold = true, isMerged = false });

            return prof;

        }

        static void Main(string[] args)
        {
            var profile = BuildProfile();

            using (var stream = new StreamWriter(@"d:\profile.xml"))
            {
                XmlSerializer xml = new XmlSerializer(typeof(EngineerProfile));
                xml.Serialize(stream, profile);
            }

        }
    }
}

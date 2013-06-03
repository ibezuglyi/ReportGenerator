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
          
            return prof;

        }

        static void Main(string[] args)
        {
            var profile = BuildProfileP();

            ReadFromFile(@"d:\profile.old", profile.OldProfile);
            ReadFromFile(@"d:\profile.new", profile.NewProfile);

            using (var stream = new StreamWriter(@"d:\profile.xml"))
            {
                XmlSerializer xml = new XmlSerializer(typeof(EngineerProfile));
                xml.Serialize(stream, profile);
            }
        }

        private static void ReadFromFile(string path, Technologies techs)
        {
            var profs = File.ReadAllLines(path);
            foreach (var p in profs)
            {
                var item = new TechnologyItem();
                if (p.EndsWith("g"))
                {
                    item.Technology = p.Substring(0, p.Length - 1);
                    item.isBold = true;
                    item.isMerged = true;
                    item.Color = 10079487;

                }
                else
                    item.Technology = p;

                item.Technology = item.Technology.Trim();

                techs.TechnologyList.Add(item);
            }
            
          
        }

        private static EngineerProfile BuildProfileP()
        {
            var profile = BuildProfile();

            using (var stream = new StreamWriter(@"d:\profile.xml"))
            {
                XmlSerializer xml = new XmlSerializer(typeof(EngineerProfile));
                xml.Serialize(stream, profile);
            }
            return profile;
        }
    }
}

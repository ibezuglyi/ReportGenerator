using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace ReportGenerator.Configuration
{

    public enum Method
    {
        //Match
        Default = 0,
        Min = 1,
        Avg = 2,
        Max = 3

    }
    [XmlRoot]
    public class Configuration
    {
        public Configuration()
        {
            Tester = new Profile();
            NetDeveloper = new Profile();
            JavaDeveloper = new Profile();
        }
        [XmlElement]
        public Profile Tester { get; set; }
        [XmlElement]
        public Profile NetDeveloper { get; set; }
        [XmlElement]
        public Profile JavaDeveloper { get; set; }
    }
    public class Profile
    {
        [XmlArray]
        public List<MappingItem> Mappings { get; set; }
        [XmlArray]
        [XmlArrayItem("KeyTechnology")]
        public List<string> KeyTechnologies { get; set; }

        public Profile()
        {
            Mappings = new List<MappingItem>();
            KeyTechnologies = new List<string>();
        }
        public void AddMappings(string from, List<string> to)
        {
            Mappings.Add(new MappingItem() { FromTechnology = from, ToTechnologies = to });
        }
        public void AddMappings(string from, params string[] to)
        {
            if (to != null)
                AddMappings(from, to.ToList());
        }
    }
    public class MappingItem
    {
        [XmlElement]
        public string FromTechnology { get; set; }
        
        [XmlArray]
        [XmlArrayItem("ToTechnology")]
        public List<string> ToTechnologies { get; set; }
        [XmlElement]
        public Method Method { get; set; }

        public MappingItem()
        {
            ToTechnologies = new List<string>();
        }
    }

}

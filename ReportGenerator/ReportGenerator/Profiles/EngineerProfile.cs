using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace ReportGenerator.Profiles
{
    [XmlRoot]
    public class EngineerProfile
    {
        [XmlElement]
        public Technologies OldProfile { get; set; }
        [XmlElement]
        public Technologies NewProfile { get; set; }

        public EngineerProfile()
        {
            OldProfile = new Technologies();
            NewProfile = new Technologies();
        }
    }

    public class Technologies
    {
        [XmlArray]
        public List<TechnologyItem> TechnologyList { get; set; }
        public Technologies()
        {
            TechnologyList = new List<TechnologyItem>();
        }

    }
    public class TechnologyItem
    {
        [XmlAttribute]
        public string  Technology { get; set; }
        [XmlAttribute]
        public int Color { get; set; }
        [XmlAttribute]
        public bool isBold { get; set; }
        [XmlAttribute]
        public bool isMerged { get; set; }
        [XmlAttribute]
        public bool IsKeyWord { get; set; }
    }
}

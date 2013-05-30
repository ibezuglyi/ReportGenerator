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
        public Header Header { get; set; }

        [XmlAttribute]
        public Profile Profile { get; set; }

        [XmlElement]
        public Technologies OldProfile { get; set; }
        [XmlElement]
        public Technologies NewProfile { get; set; }

        public EngineerProfile()
        {
            OldProfile = new Technologies();
            NewProfile = new Technologies();
            Header = new Header();
        }

        public IEnumerable<string> GetProfileKeyWords()
        {
            List<string> allTechnologies = new List<string>();
            allTechnologies.AddRange(OldProfile.TechnologyList.Where(r => r.IsKeyWord).Select(r => r.Technology));
            allTechnologies.AddRange(NewProfile.TechnologyList.Where(r => r.IsKeyWord).Select(r => r.Technology));

            return allTechnologies.Distinct();
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
        public string Technology { get; set; }
        [XmlAttribute]
        public int Color { get; set; }
        [XmlAttribute]
        public bool isBold { get; set; }
        [XmlAttribute]
        public bool isMerged { get; set; }
        [XmlAttribute]
        public bool IsKeyWord { get; set; }
        [XmlAttribute]
        public Method Method { get; set; }

        [XmlAttribute]
        public string Scale { get; set; }
        [XmlArray]
        [XmlArrayItem("MapToTechnology")]
        public List<string> MapToTechnologies { get; set; }


        public TechnologyItem()
        {
            MapToTechnologies = new List<string>();
        }
    }
    public class Header
    {
        public Header()
        {
            Scales = new List<string>();
        }
        [XmlArray]
        [XmlArrayItem("ScaleDescription")]
        public List<string> Scales { get; set; }
    }
    public enum Profile
    {
        NetDeveloper,
        JavaDeveloper,
        Tester
    }
    public enum Method
    {
        //Match
        Default = 0,
        Min = 1,
        Avg = 2,
        Max = 3

    }
}

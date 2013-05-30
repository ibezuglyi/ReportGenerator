using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace ReportGenerator.Helpers
{
    public class XmlLoader
    {

        public static Tout LoadFromXml<Tout>(string path) where Tout : class
        {
            Tout deserializedObject;
            using (var stream = new StreamReader(path))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Tout));
                deserializedObject = serializer.Deserialize(stream) as Tout;
            }

            return deserializedObject;
        }
    }
}

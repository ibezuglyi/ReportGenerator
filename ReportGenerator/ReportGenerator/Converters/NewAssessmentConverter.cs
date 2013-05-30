using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportGenerator.Converters
{
    /// <summary>
    /// converts from old format to new format
    /// </summary>
    public class NewAssessmentConverter : AssessmentConverter
    {
        protected override List<Profiles.TechnologyItem> ConvertAssessment(Assessment assessment, Profiles.EngineerProfile profile)
        {

            var technologiesList = profile.NewProfile.TechnologyList;
            var assessmentTechnologyList = assessment.GetTechnologyScales().ToDictionary(r => r.Technology, r => r.Scale);
            foreach (var technology in technologiesList)
            {
                var scales = GetScales(assessmentTechnologyList, technology.MapToTechnologies);
                string scale = GetScale(scales, technology.Method);
                technology.Scale = scale;
            }
            return technologiesList;
        }

        private string GetScale(List<string> scales, Profiles.Method method)
        {
            //process scale depends on method... but now, first not null only
            var integerScales = Array.ConvertAll(scales.Where(r => !string.IsNullOrEmpty(r)).ToArray(), int.Parse);

            return integerScales.Count() == 0 ? string.Empty : integerScales.Max().ToString();
        }

        private List<string> GetScales(Dictionary<string, string> assessmentTechnologyList, List<string> list)
        {
            var resultScales = new List<string>();
            foreach (var key in list)
            {

                if (assessmentTechnologyList.ContainsKey(key))
                    resultScales.Add(assessmentTechnologyList[key]);
            }
            return resultScales;
        }
    }
}

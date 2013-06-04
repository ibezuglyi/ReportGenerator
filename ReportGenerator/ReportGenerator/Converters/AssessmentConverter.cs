using ReportGenerator.Helpers;
using ReportGenerator.Profiles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ReportGenerator.Converters
{
    public abstract class AssessmentConverter
    {
        public abstract string ConverterName { get; }

        public List<TechnologyItem> Convert(Assessment assessment, EngineerProfile profile)
        {
            return ConvertAssessment(assessment, profile);
        }

        protected abstract List<TechnologyItem> ConvertAssessment(Assessment assessment, EngineerProfile profile);
        protected List<Profiles.TechnologyItem> ConvertProfile(Assessment assessment, List<Profiles.TechnologyItem> technologiesList)
        {

            Dictionary<string, string> assessmentTechnologyList = new Dictionary<string, string>();
            assessment.GetTechnologyScales()
                .ForEach(t =>
                {
                    if (!assessmentTechnologyList.ContainsKey(t.Technology))
                        assessmentTechnologyList.Add(t.Technology, t.Scale);
                });

            foreach (var technology in technologiesList)
            {
                var scales = GetScales(assessmentTechnologyList, technology.MapToTechnologies);
                string scale = GetScale(scales, technology.Method);
                technology.Scale = scale;
            }
            return technologiesList;
        }

        private string CalculateScale(int[] integerScales, Method method)
        {
            switch (method)
            {
                case Method.Default:
                case Method.Max:
                    return integerScales.Max().ToString();
                case Method.Avg:
                    return ((int)integerScales.Average()).ToString();
                case Method.Min:
                    return integerScales.Min().ToString();
            }
            return string.Empty;
        }
        private string GetScale(List<string> scales, Profiles.Method method)
        {
            //process scale depends on method... but now, first not null only
            var integerScales = Array.ConvertAll(scales.Where(r => !string.IsNullOrEmpty(r)).ToArray(), int.Parse);

            return integerScales.Count() == 0 ? string.Empty : CalculateScale(integerScales, method);
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

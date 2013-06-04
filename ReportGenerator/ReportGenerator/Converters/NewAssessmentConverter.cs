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
            return ConvertProfile(assessment, technologiesList);
        }

        public override string ConverterName
        {
            get { return "New format "; }
        }
    }
}

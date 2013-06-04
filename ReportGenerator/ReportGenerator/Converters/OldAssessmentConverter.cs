using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportGenerator.Converters
{
    /// <summary>
    /// converts from new format to old format
    /// </summary>
    public class OldAssessmentConverter : AssessmentConverter
    {
        protected override List<Profiles.TechnologyItem> ConvertAssessment(Assessment assessment, Profiles.EngineerProfile profile)
        {
            var technologiesList = profile.OldProfile.TechnologyList;
            return ConvertProfile(assessment, technologiesList);
        }

        public override string ConverterName
        {
            get { return "Old format "; }
        }
    }
}

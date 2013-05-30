using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportGenerator.Converters
{
    public class OldAssessmentConverter : AssessmentConverter
    {

        protected override List<Profiles.TechnologyItem> ConvertAssessment(Assessment assessment, Profiles.EngineerProfile profile)
        {
            throw new NotImplementedException();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportGenerator.Converters
{
    public class OldAssessmentConverter : AssessmentConverter
    {
        protected override Assessment ConvertAssessment(Assessment assessment)
        {
            if (assessment == null)
                return null;

            if (assessment.Version == Version.New)
                return assessment;

            Assessment oldAssessment = new Assessment();

            return oldAssessment;
        }
    }
}

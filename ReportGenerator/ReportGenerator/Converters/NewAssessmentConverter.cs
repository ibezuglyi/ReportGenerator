using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportGenerator.Converters
{
    public class NewAssessmentConverter : AssessmentConverter
    {
        protected override Assessment ConvertAssessment(Assessment assessment)
        {
            if (assessment == null)
                return null;

            if (assessment.Version == Version.New)
                return assessment;
//            var profile = DetectProfile(assessment)
            Assessment newAssessment = new Assessment();

            foreach (var skill in assessment.TechnicalSkills)
            { 
                
            }
            return newAssessment;
        }
    }
}

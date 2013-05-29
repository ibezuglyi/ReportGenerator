using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ReportGenerator
{
    public enum Version
    {
        Old = 0,
        New = 1
    }
    public class Assessment
    {
        static readonly int whiteColor = ColorTranslator.ToOle(Color.White);
        const int maxRow = 255;
        public Version Version { get; set; }
        readonly string[] Header = { "Scale (0-4)", "0 - No experience.",
                                "1 - Beginner (is able to perform simple task under supervision)",
                                "2 - Autonomous (is able to perform task without supervision but may need support from expert)",
                                "3 - Expert (is able to perform task without supervision, mentor/coach for others)",
                                "4 - Senior Developer (guru)"};
        public IList<TechnicalSkill> TechnicalSkills { get; set; }


        public Assessment()
        {
            TechnicalSkills = new List<TechnicalSkill>();
        }

        public static Assessment Build(Microsoft.Office.Interop.Excel.Worksheet ActiveSheet)
        {
            int startRow = GetStartRowNumber(ActiveSheet);
            if (startRow == 0)
                //technical area not found
                return null;

            string column = GetColumnLetter(ActiveSheet, startRow);
            if (column == null)
                //scale column not found
                return null;

            Color groupColor = Color.FromArgb(255, 204, 255, 204);
            var oleGroupColor = ColorTranslator.ToOle(groupColor);

            Assessment assessment = new Assessment();
            int currentRow = startRow;
            currentRow = FindNewGroup(currentRow, ActiveSheet, oleGroupColor);
            IList<TechnicalSkill> skillGroups = BuildTechnicalSkill(currentRow, column, oleGroupColor, ActiveSheet);
            assessment.TechnicalSkills = skillGroups;
            return assessment;
        }

        private static IList<TechnicalSkill> BuildTechnicalSkill(int currentRow, string column, int groupColor, Microsoft.Office.Interop.Excel.Worksheet ActiveSheet)
        {
            List<TechnicalSkill> skillList = new List<TechnicalSkill>();
            while (true)
            {
                int nextGroupRowNumber = FindNewGroup(currentRow+1, ActiveSheet, groupColor);
                if (nextGroupRowNumber == currentRow + 1)
                    break;

                TechnicalSkill skill = new TechnicalSkill();
                skill.Technology = GetTechnologyName(currentRow, ActiveSheet);
                currentRow++;
                skill.Technologies = GetTechnologies(currentRow, nextGroupRowNumber, column, ActiveSheet);
                currentRow++;
                skillList.Add(skill);
                currentRow = nextGroupRowNumber;
            }

            return skillList;
        }

        private static IList<TechnologyGroupItem> GetTechnologies(int currentRow, int nextGroupRowNumber, string column, Microsoft.Office.Interop.Excel.Worksheet ActiveSheet)
        {
            var technologyGroupList = new List<TechnologyGroupItem>();
            for (int row = currentRow; row < nextGroupRowNumber; row++)
            {
                TechnologyGroupItem item = new TechnologyGroupItem();
                item.Technology = GetAValue(row, ActiveSheet);
                string scale = GetValue(row, column, ActiveSheet);
                item.Scale = string.IsNullOrEmpty(scale) ? (byte?)null : byte.Parse(scale);
                technologyGroupList.Add(item);
            }
            return technologyGroupList;
        }

        private static string GetTechnologyName(int currentRow, Microsoft.Office.Interop.Excel.Worksheet ActiveSheet)
        {
            return GetAValue(currentRow, ActiveSheet);
        }

        private static string GetAValue(int currentRow, Microsoft.Office.Interop.Excel.Worksheet ActiveSheet)
        {
            return GetValue(currentRow, "A", ActiveSheet);

        }
        private static string GetValue(int currentRow, string currentColumn, Microsoft.Office.Interop.Excel.Worksheet ActiveSheet)
        {
            var cell = string.Format("{0}{1}", currentColumn, currentRow);
            return ActiveSheet.Range[cell].Value2;
        }

        private static int FindNewGroup(int currentRow, Microsoft.Office.Interop.Excel.Worksheet ActiveSheet, int oleGroupColor)
        {
            while (currentRow < maxRow)
            {
                var cell = string.Format("A{0}", currentRow);
                int cellColor = (int)ActiveSheet.Range[cell].Interior.Color;

                if (oleGroupColor == cellColor)
                    //new group found
                    return currentRow;
                else
                    if (string.IsNullOrEmpty(ActiveSheet.Range[cell].Value2) && cellColor == whiteColor)
                        return currentRow;

                currentRow++;
            }
            return 0;
        }

        private static string GetColumnLetter(Microsoft.Office.Interop.Excel.Worksheet ActiveSheet, int startRow)
        {
            char columnLetter = 'B';
            while (columnLetter < 'Z')
            {
                string cell = string.Format("{0}{1}", columnLetter, startRow);
                string value = ActiveSheet.Range[cell].Value2;
                if (!string.IsNullOrEmpty(value) &&
                    value.Contains("Scale"))
                    return columnLetter.ToString();

                columnLetter++;
            }
            return null;


        }

        private static int GetStartRowNumber(Microsoft.Office.Interop.Excel.Worksheet ActiveSheet)
        {
            int currentRow = 1;
            int startRow = 0;
            string cell = null;
            while (startRow == 0 && currentRow < maxRow)
            {
                cell = string.Format("A{0}", currentRow);
                string cellValue = ActiveSheet.Range[cell].Value2;
                if (cellValue == "Technical Area")
                    startRow = currentRow;
                else
                    currentRow++;
            }
            return startRow;
        }
    }
    public class TechnicalSkill
    {
        public string Technology { get; set; }
        public IList<TechnologyGroupItem> Technologies { get; set; }

        public TechnicalSkill()
        {
            Technologies = new List<TechnologyGroupItem>();
        }
    }
    public class TechnologyGroupItem
    {
        public string Technology { get; set; }
        public byte? Scale { get; set; }
    }
}

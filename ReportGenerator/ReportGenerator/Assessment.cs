using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Drawing;
using ReportGenerator.Helpers;

namespace ReportGenerator
{
    public enum Version
    {
        Old = 0,
        New = 1
    }
    public class Assessment
    {
        public static readonly int OleWhiteColor = ColorTranslator.ToOle(Color.White);
        public static readonly int OleGroupColor = ColorTranslator.ToOle(Color.FromArgb(255, 204, 255, 204));
        public static readonly int OleHeaderColor = 10079487;

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
        public IList<string> GetAllTechnologies()
        {
            List<string> technologyList = new List<string>();
            foreach (var item in TechnicalSkills)
            {
                technologyList.AddRange(
                 item.Technologies.Select(r => r.Technology));
            }
            return technologyList;
        }

        public List<TechnologyGroupItem> GetTechnologyScales()
        {
            List<TechnologyGroupItem> items = new List<TechnologyGroupItem>();
            foreach (var technicalSkill in TechnicalSkills)
            {
                TechnologyGroupItem item = new TechnologyGroupItem() { Technology = technicalSkill.Technology, Scale = string.Empty };
                items.Add(item);
                foreach (var technology in technicalSkill.Technologies)
                    items.Add(new TechnologyGroupItem() { Scale = technology.Scale, Technology = technology.Technology });
            }
            return items;
        }
        public static Assessment Build(Microsoft.Office.Interop.Excel.Worksheet ActiveSheet, IEnumerable<string> technologyGroups)
        {
            int startRow = GetStartRowNumber(ActiveSheet, "Technical Area", "Technical Knowledge", "General testing knowledge");
            if (startRow == 0)
                //technical area not found
                return null;

            string column = GetColumnLetter(ActiveSheet, startRow);
            if (column == null)
                //scale column not found
                return null;

            Assessment assessment = new Assessment();
            int currentRow = startRow;
            currentRow = FindNewGroup(currentRow, ActiveSheet, technologyGroups);
            IList<TechnicalSkill> skillGroups = BuildTechnicalSkill(currentRow, column, technologyGroups, ActiveSheet);
            assessment.TechnicalSkills = skillGroups;
            return assessment;
        }

        private static IList<TechnicalSkill> BuildTechnicalSkill(int currentRow, string column, IEnumerable<string> technologyGroups, Microsoft.Office.Interop.Excel.Worksheet activeSheet)
        {
            ExcelWorker worker = new ExcelWorker(activeSheet);
            List<TechnicalSkill> skillList = new List<TechnicalSkill>();
            while (true)
            {
                int nextGroupRowNumber = FindNewGroup(currentRow + 1, activeSheet, technologyGroups);
                if (nextGroupRowNumber == currentRow + 1)
                    break;

                TechnicalSkill skill = new TechnicalSkill();
                skill.Technology = GetTechnologyName(currentRow, worker);
                currentRow++;
                skill.Technologies = GetTechnologies(currentRow, nextGroupRowNumber, column, worker);
                currentRow++;
                skillList.Add(skill);
                currentRow = nextGroupRowNumber;
            }

            return skillList;
        }

        private static IList<TechnologyGroupItem> GetTechnologies(int currentRow, int nextGroupRowNumber, string column, ExcelWorker worker)
        {
            var technologyGroupList = new List<TechnologyGroupItem>();
            for (int row = currentRow; row < nextGroupRowNumber; row++)
            {
                TechnologyGroupItem item = new TechnologyGroupItem();
                item.Technology = worker.GetAValue(row);
                item.Scale = worker.GetValue(row, column);
                technologyGroupList.Add(item);
            }
            return technologyGroupList;
        }

        private static string GetTechnologyName(int currentRow, ExcelWorker worker)
        {
            return worker.GetAValue(currentRow);
        }
        private static int FindNewGroup(int currentRow, Microsoft.Office.Interop.Excel.Worksheet ActiveSheet, IEnumerable<string> technologyGroups)
        {
            while (currentRow < maxRow)
            {
                var cell = string.Format("A{0}", currentRow);
                int cellColor = (int)ActiveSheet.Range[cell].Interior.Color;
                //int fontsize = (int)ActiveSheet.Range[cell].Style.Font.Size;
                //bool isBold = (bool)ActiveSheet.Range[cell].Style.Font.Bold;
                //if ((oleGroupColor == cellColor) || //new format condition
                //    (fontsize == 12 && isBold))     //old format condition

                string value = Convert.ToString(ActiveSheet.Range[cell].Value2);
                if (technologyGroups.Contains(value))
                    //new group found
                    return currentRow;
                else
                    if (string.IsNullOrEmpty(ActiveSheet.Range[cell].Value2) && cellColor == OleWhiteColor)
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
            return "B";


        }

        private static int GetStartRowNumber(Microsoft.Office.Interop.Excel.Worksheet ActiveSheet, params string[] headerTexts)
        {
            int currentRow = 1;
            int startRow = 0;
            string cell = null;
            while (startRow == 0 && currentRow < maxRow)
            {
                cell = string.Format("A{0}", currentRow);
                string cellValue = ActiveSheet.Range[cell].Value2;
                if (headerTexts.Contains(cellValue))
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
        public string Scale { get; set; }
    }
}

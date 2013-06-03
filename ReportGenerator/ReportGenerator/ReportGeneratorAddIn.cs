using ReportGenerator.Converters;
using ReportGenerator.Helpers;
using ReportGenerator.Profiles;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportGenerator
{
    public partial class ReportGeneratorAddIn
    {
        private Excel.Worksheet sheet;
        private Excel.Worksheet ActiveSheet
        {
            get
            {
                if (sheet != null)
                {
                    return sheet;
                }
                else
                {
                    return Application.ActiveSheet;
                }
            }

            set
            {
                sheet = value;
            }
        }
        private Excel.Workbook ActiveWorkbook { get; set; }
        private Excel.Sheets Sheets { get; set; }
        private Excel.Workbooks Workbooks { get; set; }
        private Excel.Workbook ThisWorkbook { get; set; }

        public Dictionary<string, EngineerProfile> ProfilesDictionary = new Dictionary<string, EngineerProfile>();
        public EngineerProfile TesterProfile { get; set; }
        public EngineerProfile NetDeveloperProfile { get; set; }
        public EngineerProfile JavaDeveloperProfile { get; set; }

        #region Event_Handlers
        void Application_SheetActivate(object Sh)
        {
            ActiveSheet = Application.ActiveSheet;
        }

        void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            ActiveWorkbook = Application.ActiveWorkbook;
            ActiveSheet = Application.ActiveSheet;
            Sheets = Application.ActiveWorkbook.Sheets;
        }

        void Application_WorkbookNewSheet(Excel.Workbook Wb, object Sh)
        {
            Sheets = Application.ActiveWorkbook.Sheets;
            Workbooks = Application.Workbooks;
        }

        void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            ActiveWorkbook = Application.ActiveWorkbook;
            ActiveSheet = Application.ActiveSheet;
            Sheets = Application.ActiveWorkbook.Sheets;
            Workbooks = Application.Workbooks;
        }
        #endregion

        private void ReportGeneratorAddIn_Startup(object sender, System.EventArgs e)
        {
            App.ReportGeneratorAddIn = this;
            Application.WorkbookActivate += Application_WorkbookActivate;
            Application.SheetActivate += Application_SheetActivate;
            Application.WorkbookOpen += Application_WorkbookOpen;
            Application.WorkbookNewSheet += Application_WorkbookNewSheet;
            ActiveWorkbook = Application.ActiveWorkbook;
            ActiveSheet = Application.ActiveSheet;

        }

        private void ReportGeneratorAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Application.WorkbookActivate -= Application_WorkbookActivate;
            Application.SheetActivate -= Application_SheetActivate;
            Application.WorkbookOpen -= Application_WorkbookOpen;
            Application.WorkbookNewSheet -= Application_WorkbookNewSheet;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ReportGeneratorAddIn_Startup);
            this.Shutdown += new System.EventHandler(ReportGeneratorAddIn_Shutdown);
        }

        #endregion

        public void ConvertAssessment(AssessmentConverter converter)
        {
            Assessment oldAssessment = Assessment.Build(ActiveSheet);
            //var newAssessment = converter.Convert(oldAssessment, ReportConfiguration.Instance.ConfigurationFilePath);

            EngineerProfile profile = LoadEngineerProfile(oldAssessment);
            if (profile == null)
            {
                System.Windows.Forms.MessageBox.Show("Unable to detect a profile of assessment, ensure please: 1) profile configuration has keywords selected, 2) assessment has keywords.");
                return;
            }
            var excelRows = converter.Convert(oldAssessment, profile);

            var newWorkBook = Workbooks.Add();
            var activeSheet = newWorkBook.ActiveSheet as Excel.Worksheet;
            //write header
            var worker = new ExcelWorker(activeSheet);
            int row = 2;
            foreach (var header in profile.Header.Scales)
            {
                worker.SetAValue(row++, header);
            }
            //2 rows separation 
            row += 2;
            worker.SetAValue(row, "Technical Area").SetBold(true).SetColor(Assessment.OleHeaderColor).SetWidth(60);
            worker.SetValue(row, "B", "Scale 0-4").SetBold(true).SetColor(Assessment.OleHeaderColor);

            row++;
            foreach (var item in excelRows)
            {
                worker.SetAValue(row, item.Technology).SetColor(item.Color).SetBold(item.isBold);
                worker.SetValue(row, "B", item.Scale).SetColor(item.Color).SetBold(item.isBold);
                row++;
            }

            newWorkBook.SaveAs("a.xls", Excel.XlFileFormat.xlOpenXMLWorkbook);
        }



        private EngineerProfile LoadEngineerProfile(Assessment assessment)
        {
            ProfilesDictionary.Clear();
            ProfilesDictionary.Add("tester", LoadProfile("tester"));
            ProfilesDictionary.Add("netdeveloper", LoadProfile("netdeveloper"));
            ProfilesDictionary.Add("javadeveloper", LoadProfile("javadeveloper"));
            return DetectProfile(ProfilesDictionary, assessment);
        }

        private EngineerProfile LoadProfile(string pattern)
        {
            EngineerProfile profile = null;
            string path = ReportConfiguration.Instance.ConfigurationProfileDirectory;
            pattern = string.Format("{0}.profile", pattern);
            var configFile = Directory.GetFiles(path, pattern, SearchOption.TopDirectoryOnly).SingleOrDefault();
            if (string.IsNullOrEmpty(configFile))
                System.Windows.Forms.MessageBox.Show(string.Format("Unable to load {0} config file, ensure file exists in {1} ", pattern, path), "Error");
            else
            {
                profile = XmlLoader.LoadFromXml<EngineerProfile>(configFile);
            }
            return profile;
        }
        private EngineerProfile DetectProfile(Dictionary<string, EngineerProfile> profilesDictionary, Assessment oldAssessment)
        {
            var allTechnologies = oldAssessment.GetAllTechnologies();
            foreach (var profile in profilesDictionary.Values)
            {
                if (profile != null)
                {
                    var keywords = profile.GetProfileKeyWords();
                    if (keywords.Count() > 0 && keywords.All(r => allTechnologies.Contains(r)))
                        return profile;
                }
            }

            return null;

        }

    }
}

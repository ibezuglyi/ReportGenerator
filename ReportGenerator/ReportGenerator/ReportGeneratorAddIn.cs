using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;
using ReportGenerator.Converters;

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
            var newAssessment = converter.Convert(oldAssessment, ReportConfiguration.Instance.ConfigurationFilePath);
        }

        
    }
}

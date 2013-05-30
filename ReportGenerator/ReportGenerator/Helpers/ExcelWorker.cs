using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportGenerator.Helpers
{
    public class ExcelWorker
    {
        public ExcelWorker(Microsoft.Office.Interop.Excel.Worksheet activeSheet)
        {
            this.ActiveSheet = activeSheet;
        }

        public Microsoft.Office.Interop.Excel.Worksheet ActiveSheet { get; set; }


        public string GetAValue(int currentRow)
        {
            return GetValue(currentRow, "A");
        }
        public void SetAValue(int currentRow, string value)
        {
            SetValue(currentRow, "A", value);
        }

        public void SetValue(int currentRow, string currentColumn, string value)
        {
            var cell = GetCell(currentRow, currentColumn);
            SetValue(cell, value);
        }
        public string GetValue(int currentRow, string currentColumn)
        {
            var cell = GetCell(currentRow, currentColumn);
            return GetValue(cell);
        }

        private static string GetCell(int currentRow, string currentColumn)
        {
            return string.Format("{0}{1}", currentColumn, currentRow);
        }

        private void SetValue(string cell, string value)
        {
            ActiveSheet.Range[cell].Value2 = value;
        }
        private string GetValue(string cell)
        {
            return ActiveSheet.Range[cell].Value2;
        }
    }
}

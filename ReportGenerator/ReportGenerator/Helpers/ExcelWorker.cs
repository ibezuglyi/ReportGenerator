using Microsoft.Office.Interop.Excel;
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
        public Range SetAValue(int currentRow, string value)
        {
            return SetValue(currentRow, "A", value);
        }

        public Range SetValue(int currentRow, string currentColumn, string value)
        {
            var cell = GetCell(currentRow, currentColumn);
            return SetValue(cell, value);
        }
        public string GetValue(int currentRow, string currentColumn)
        {
            var cell = GetCell(currentRow, currentColumn);
            var value = GetValue(cell);
            return string.IsNullOrEmpty(value) ? value : value.Trim();
        }

        private static string GetCell(int currentRow, string currentColumn)
        {
            return string.Format("{0}{1}", currentColumn, currentRow);
        }

        private Range SetValue(string cell, string value)
        {
            ActiveSheet.Range[cell].Value2 = value;
            return ActiveSheet.Range[cell];
        }
        private string GetValue(string cell)
        {
            return Convert.ToString(ActiveSheet.Range[cell].Value2);
        }
    }

    public static class ExcelHelper
    {
        public static Range SetColor(this Range range, int color)
        {
            if (color != 0)
                range.Interior.Color = color;

            return range;
        }
        public static Range SetHeight(this Range range, int height)
        {
            range.EntireRow.RowHeight = height;
            return range;
        }
        public static Range SetWidth(this Range range, int width)
        {
            range.EntireColumn.ColumnWidth = width;
            return range;
        }
        public static Range SetAlignment(this Range range, Microsoft.Office.Interop.Excel.XlHAlign alignment)
        {
            range.Style.HorizontalAlignment = alignment;
            return range;
        }
        public static Range SetBold(this Range range, bool isBold)
        {
            range.Style.Font.Bold = isBold;
            return range;
        }
    }
}

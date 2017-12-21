using System;
using ExportToExcel.StylesheetProvider;

namespace ExportToExcel.Models
{
    public class ExcelCell
    {
        public string Value { get; set; }
        public ExcelSheetStyleIndex StyleIndex { get; set; }
        public Uri Uri { get; set; }

        public ExcelCell(string value, 
            ExcelSheetStyleIndex styleIndex = ExcelSheetStyleIndex.Default,
            Uri uri = null)
        {
            Value = value;
            StyleIndex = styleIndex;
            Uri = uri;
        }
    }
}
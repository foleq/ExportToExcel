using ExportToExcel.StylesheetProvider;

namespace ExportToExcel.Models
{
    public class ExcelCell
    {
        public string Value { get; set; }
        public ExcelSheetStyleIndex StyleIndex { get; set; }

        public ExcelCell(string value, ExcelSheetStyleIndex styleIndex = ExcelSheetStyleIndex.Default)
        {
            Value = value;
            StyleIndex = styleIndex;
        }
    }
}
using System;
using ExportToExcel.Models;

namespace ExportToExcel.Providers
{
    public interface IExcelColumnBuilder
    {
        ExcelColumn GetColumn(int maxNumberOfCharactersInColumn, uint columnNumber);
        ExcelColumn GetColumn(int maxNumberOfCharactersInColumn, uint columnNumberStart, uint columnNumberEnd);
    }

    public class ExcelColumnProvider : IExcelColumnBuilder
    {
        private const double MaxWidthOfFont = 7;

        public ExcelColumn GetColumn(int maxNumberOfCharactersInColumn, uint columnNumber)
        {
            return GetColumn(maxNumberOfCharactersInColumn, columnNumber, columnNumber);
        }

        public ExcelColumn GetColumn(int maxNumberOfCharactersInColumn, uint columnNumberStart, uint columnNumberEnd)
        {
            // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column(v=office.14).aspx
            // width = Truncate([{Nformat4Decimal of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
            var width = Math.Truncate((maxNumberOfCharactersInColumn * MaxWidthOfFont + 5) / MaxWidthOfFont * 256) / 256;

            return new ExcelColumn()
            {
                Width = width,
                ColumnNumberStart = columnNumberStart,
                ColumnNumberEnd = columnNumberEnd,
            };
        }
    }
}
using System;

namespace ExportToExcel.Providers
{
    public interface IExcelCellNameProvider
    {
        string GetCellName(int columnNumber, int rowNumber);
    }

    public class ExcelCellNameProvider: IExcelCellNameProvider
    {
        public string GetCellName(int columnNumber, int rowNumber)
        {
            if (columnNumber <= 0 || rowNumber <= 0)
            {
                throw new ArgumentException("Column and row number should be greater than 0.");
            }
            var dividend = columnNumber;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName + rowNumber;
        }
    }
}
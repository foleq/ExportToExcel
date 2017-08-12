using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.Models;
using ExportToExcel.StylesheetProvider;

namespace ExportToExcel.Factories
{
    //TODO: Could be internal (public for tests)
    public interface IExcelCellFactory
    {
        Cell GetCell(ExcelCell excelCell);
    }

    public class ExcelCellFactory : IExcelCellFactory
    {
        private readonly IExcelStylesheetProvider _stylesheetProvider;

        public ExcelCellFactory(IExcelStylesheetProvider stylesheetProvider)
        {
            _stylesheetProvider = stylesheetProvider;
        }

        public Cell GetCell(ExcelCell excelCell)
        {
            if (string.IsNullOrEmpty(excelCell?.Value))
            {
                return new Cell();
            }

            var dataType = GetDataTypeAndUpdateCellValueIfNecessary(excelCell);
            return new Cell()
            {
                DataType = dataType,
                StyleIndex = _stylesheetProvider.GetSheetStyleIndex(excelCell.StyleIndex),
                CellValue = new CellValue(excelCell.Value),
            };
        }

        private static CellValues GetDataTypeAndUpdateCellValueIfNecessary(ExcelCell excelCell)
        {
            if (IsNumericType(excelCell.StyleIndex))
            {
                double numericValue;
                if (double.TryParse(excelCell.Value, out numericValue))
                {
                    excelCell.Value = numericValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    return CellValues.Number;
                }
                excelCell.StyleIndex = ExcelSheetStyleIndex.Default;
            }
            return CellValues.String;
        }

        private static bool IsNumericType(ExcelSheetStyleIndex styleIndex)
        {
            return styleIndex == ExcelSheetStyleIndex.Number;
        }
    }
}
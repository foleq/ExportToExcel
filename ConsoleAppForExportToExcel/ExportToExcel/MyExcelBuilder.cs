using System;
using ExportToExcel.Builders;
using ExportToExcel.Models;
using ExportToExcel.StylesheetProvider;

namespace ConsoleAppForExportToExcel.ExportToExcel
{
    public interface IMyExcelBuilder
    {
        byte[] BuildExcelFile(int numberOfWorksheet, int numberOfColumns, int numberOfRows);
    }

    public class MyExcelBuilder : IMyExcelBuilder
    {
        private readonly IExcelBuilder _excelBuilder;

        public MyExcelBuilder(IExcelBuilder excelBuilder)
        {
            _excelBuilder = excelBuilder;
        }

        public byte[] BuildExcelFile(int numberOfWorksheet, int numberOfColumns, int numberOfRows)
        {
            for (var worksheetIndex = 0; worksheetIndex < numberOfWorksheet; worksheetIndex++)
            {
                var worksheetName = $"worksheet_{worksheetIndex}";

                var cells = new ExcelCell[numberOfColumns];
                for (var columnIndex = 0; columnIndex < numberOfColumns; columnIndex++)
                {
                    cells[columnIndex] = new ExcelCell($"cell_name_{columnIndex}", ExcelSheetStyleIndex.Bold);
                }
                _excelBuilder.AddRowToWorksheet(worksheetName, cells);

                for (var rowIndex = 0; rowIndex < numberOfRows; rowIndex++)
                {
                    cells = new ExcelCell[numberOfColumns];
                    for (var columnIndex = 0; columnIndex < numberOfColumns; columnIndex++)
                    {
                        Uri uri = null;
                        switch (columnIndex)
                        {
                            case 0:
                                uri = new Uri("http://wykop.pl");
                                break;
                            case 1:
                                uri = new Uri("http://wp.pl");
                                break;
                        }
                        cells[columnIndex] = new ExcelCell($"cell_value_{rowIndex}_{columnIndex}", ExcelSheetStyleIndex.Default, uri);
                    }
                    _excelBuilder.AddRowToWorksheet(worksheetName, cells);
                }
            }

            return _excelBuilder.FinishAndGetExcel();
        }
    }
}
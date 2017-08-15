using ExportToExcel.Builders;
using ExportToExcel.Factories;
using ExportToExcel.StylesheetProvider;

namespace ConsoleAppForExportToExcel.ExportToExcel
{
    public interface IMyExcelBuilderFactory
    {
        MyExcelBuilder CreateMyExcelBuilder();
    }

    public class MyExcelBuilderFactory : IMyExcelBuilderFactory
    {
        public MyExcelBuilder CreateMyExcelBuilder()
        {
            var excelStylesheetProvider = new ExcelStylesheetProvider(
                new ExcelStylesheetNumberingFormatProvider(), 
                new ExcelStylesheetFontProvider(), 
                new ExcelStylesheetFillProvider(), 
                new ExcelStylesheetBorderProvider()
                );
            var excelCellFactory = new ExcelCellFactory(excelStylesheetProvider);
            var excelBuilder = new ExcelBuilder(excelStylesheetProvider, excelCellFactory);
            return new MyExcelBuilder(excelBuilder);
        }
    }
}
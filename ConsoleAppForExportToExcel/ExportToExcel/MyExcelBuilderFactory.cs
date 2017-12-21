using ExportToExcel.Builders;
using ExportToExcel.Factories;
using ExportToExcel.Providers;
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
            var excelCellNameProvider = new ExcelCellNameProvider();
            var excelBuilder = new ExcelBuilder(excelStylesheetProvider, excelCellFactory, excelCellNameProvider);
            return new MyExcelBuilder(excelBuilder);
        }
    }
}
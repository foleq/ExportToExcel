using System.Linq;
using developwithpassion.specifications.rhinomocks;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.StylesheetProvider;
using Machine.Specifications;
using Rhino.Mocks;

namespace ExportToExcel.Tests.StylesheetProvider
{
    [Subject(typeof(ExcelStylesheetProvider))]
    internal class ExcelStylesheetProviderSpecs : Observes<ExcelStylesheetProvider>
    {
    }

    internal class When_getting_stylesheet : ExcelStylesheetProviderSpecs
    {
        private Establish context = () =>
        {
            var numberingFormatProvider = depends.on<IExcelStylesheetNumberingFormatProvider>();
            numberingFormatProvider.Stub(x => x.GetNumberingFormats()).Return(_numberingFormats);
            numberingFormatProvider.Stub(x => x.GetNumberFormatIndex(ExcelSheetNumberingFormatIndex.Nformat4Decimal))
                .Return(100);
            var fontProvider = depends.on<IExcelStylesheetFontProvider>();
            fontProvider.Stub(x => x.GetFonts()).Return(_fonts);
            fontProvider.Stub(x => x.GetFontIndex(ExcelStylesheetFontIndex.Default))
                .Return(100);
            fontProvider.Stub(x => x.GetFontIndex(ExcelStylesheetFontIndex.Bold))
                .Return(101);
            var fillProvider = depends.on<IExcelStylesheetFillProvider>();
            fillProvider.Stub(x => x.GetFills()).Return(_fills);
            fillProvider.Stub(x => x.GetFillIndex(ExcelStylesheetFillIndex.Default))
                .Return(100);
            var borderProvider = depends.on<IExcelStylesheetBorderProvider>();
            borderProvider.Stub(x => x.GetBorders()).Return(_borders);
            borderProvider.Stub(x => x.GetBorderIndex(ExcelSheetBorderIndex.Default))
                .Return(100);
        };

        Because of = () =>  _resultStylesheet = sut.GetStylesheet();

        It should_have_3_cell_formats = () =>
            _resultStylesheet.CellFormats.ChildElements.Count.ShouldEqual(3);

        It should_have_properly_formatted_1st_cell_format = () =>
        {
            var cellFormat = _resultStylesheet.CellFormats.ChildElements.ToArray()[0] as CellFormat;
            cellFormat.FontId.Value.ShouldEqual<uint>(100);
            cellFormat.FillId.Value.ShouldEqual<uint>(100);
            cellFormat.BorderId.Value.ShouldEqual<uint>(100);
        };

        It should_have_properly_formatted_2nd_cell_format = () =>
        {
            var cellFormat = _resultStylesheet.CellFormats.ChildElements.ToArray()[1] as CellFormat;
            cellFormat.FontId.Value.ShouldEqual<uint>(101);
            cellFormat.ApplyFont.Value.ShouldBeTrue();
        };

        It should_have_properly_formatted_3rd_cell_format = () =>
        {
            var cellFormat = _resultStylesheet.CellFormats.ChildElements.ToArray()[2] as CellFormat;
            cellFormat.NumberFormatId.Value.ShouldEqual<uint>(100);
            cellFormat.ApplyNumberFormat.Value.ShouldBeTrue();
        };

        It should_have_proper_numberFormats = () => 
            _resultStylesheet.NumberingFormats.ShouldEqual(_numberingFormats);

        It should_have_proper_fonts = () =>
            _resultStylesheet.Fonts.ShouldEqual(_fonts);

        It should_have_proper_fills = () =>
            _resultStylesheet.Fills.ShouldEqual(_fills);

        It should_have_proper_borders = () =>
            _resultStylesheet.Borders.ShouldEqual(_borders);

        private static readonly NumberingFormats _numberingFormats = new NumberingFormats();
        private static readonly Fonts _fonts = new Fonts();
        private static readonly Fills _fills = new Fills();
        private static readonly Borders _borders = new Borders();
        private static Stylesheet _resultStylesheet;
    }

    internal class When_getting_default_stylesheet_index : ExcelStylesheetProviderSpecs
    {
        Because of = () => _resultIndex = sut.GetSheetStyleIndex(ExcelSheetStyleIndex.Default);

        It should_return_proper_index = () => _resultIndex.ShouldEqual<uint>(0);

        private static uint _resultIndex;
    }

    internal class When_getting_bold_stylesheet_index : ExcelStylesheetProviderSpecs
    {
        Because of = () => _resultIndex = sut.GetSheetStyleIndex(ExcelSheetStyleIndex.Bold);

        It should_return_proper_index = () => _resultIndex.ShouldEqual<uint>(1);

        private static uint _resultIndex;
    }

    internal class When_getting_number_stylesheet_index : ExcelStylesheetProviderSpecs
    {
        Because of = () => _resultIndex = sut.GetSheetStyleIndex(ExcelSheetStyleIndex.Number);

        It should_return_proper_index = () => _resultIndex.ShouldEqual<uint>(2);

        private static uint _resultIndex;
    }
}
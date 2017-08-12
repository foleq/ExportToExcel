using developwithpassion.specifications.rhinomocks;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.StylesheetProvider;
using Machine.Specifications;

namespace ExportToExcel.Tests.StylesheetProvider
{
    [Subject(typeof(ExcelStylesheetFontProvider))]
    internal class ExcelStylesheetFontProviderSpecs : Observes<ExcelStylesheetFontProvider>
    {
    }

    internal class When_getting_fonts : ExcelStylesheetFontProviderSpecs
    {
        Because of = () =>  _result = sut.GetFonts();

        It should_have_one_font = () => _result.ChildElements.Count.ShouldEqual(2);

        private static Fonts _result;
    }

    internal class When_getting_default_font_index : ExcelStylesheetFontProviderSpecs
    {
        Because of = () => _resultIndex = sut.GetFontIndex(ExcelStylesheetFontIndex.Default);

        It should_return_proper_index = () => _resultIndex.ShouldEqual<uint>(0);

        private static uint _resultIndex;
    }

    internal class When_getting_bold_font_index : ExcelStylesheetFontProviderSpecs
    {
        Because of = () => _resultIndex = sut.GetFontIndex(ExcelStylesheetFontIndex.Bold);

        It should_return_proper_index = () => _resultIndex.ShouldEqual<uint>(1);

        private static uint _resultIndex;
    }
}
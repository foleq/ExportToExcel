using developwithpassion.specifications.rhinomocks;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.StylesheetProvider;
using Machine.Specifications;

namespace ExportToExcel.Tests.StylesheetProvider
{
    [Subject(typeof(ExcelStylesheetBorderProvider))]
    internal class ExcelStylesheetBorderProviderSpecs : Observes<ExcelStylesheetBorderProvider>
    {
    }

    internal class When_getting_borders : ExcelStylesheetBorderProviderSpecs
    {
        Because of = () =>  _result = sut.GetBorders();

        It should_have_one_border = () => _result.ChildElements.Count.ShouldEqual(1);

        private static Borders _result;
    }

    internal class When_getting_default_border_index : ExcelStylesheetBorderProviderSpecs
    {
        Because of = () => _borderIndex = sut.GetBorderIndex(ExcelSheetBorderIndex.Default);

        It should_return_proper_index = () => _borderIndex.ShouldEqual<uint>(0);

        private static uint _borderIndex;
    }
}
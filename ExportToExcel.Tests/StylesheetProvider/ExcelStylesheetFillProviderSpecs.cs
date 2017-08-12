using developwithpassion.specifications.rhinomocks;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.StylesheetProvider;
using Machine.Specifications;

namespace ExportToExcel.Tests.StylesheetProvider
{
    [Subject(typeof(ExcelStylesheetFillProvider))]
    internal class ExcelStylesheetFillProviderSpecs : Observes<ExcelStylesheetFillProvider>
    {
    }

    internal class When_getting_fills : ExcelStylesheetFillProviderSpecs
    {
        Because of = () =>  _result = sut.GetFills();

        It should_have_one_fill = () => _result.ChildElements.Count.ShouldEqual(1);

        private static Fills _result;
    }

    internal class When_getting_default_fill_index : ExcelStylesheetFillProviderSpecs
    {
        Because of = () => _resultIndex = sut.GetFillIndex(ExcelStylesheetFillIndex.Default);

        It should_return_proper_index = () => _resultIndex.ShouldEqual<uint>(0);

        private static uint _resultIndex;
    }
}
using developwithpassion.specifications.rhinomocks;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.StylesheetProvider;
using Machine.Specifications;

namespace ExportToExcel.Tests.StylesheetProvider
{
    [Subject(typeof(ExcelStylesheetNumberingFormatProvider))]
    internal class ExcelStylesheetNumberingFormatProviderSpecs : Observes<ExcelStylesheetNumberingFormatProvider>
    {
    }

    internal class When_getting_numberingFormats : ExcelStylesheetNumberingFormatProviderSpecs
    {
        Because of = () =>  _result = sut.GetNumberingFormats();

        It should_have_one_numberingFormat = () => _result.ChildElements.Count.ShouldEqual(1);

        private static NumberingFormats _result;
    }

    internal class When_getting_Nformat4Decimal_numberingFormat_index : ExcelStylesheetNumberingFormatProviderSpecs
    {
        Because of = () => _resultIndex = sut.GetNumberFormatIndex(ExcelSheetNumberingFormatIndex.Nformat4Decimal);

        It should_return_proper_index = () => _resultIndex.ShouldEqual<uint>(0);

        private static uint _resultIndex;
    }
}
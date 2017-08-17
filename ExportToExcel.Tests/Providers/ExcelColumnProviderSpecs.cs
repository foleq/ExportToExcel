using developwithpassion.specifications.rhinomocks;
using ExportToExcel.Models;
using ExportToExcel.Providers;
using Machine.Specifications;

namespace ExportToExcel.Tests.Providers
{
    [Subject(typeof(ExcelColumnProvider))]
    internal abstract class ExcelColumnProviderSpecs : Observes<ExcelColumnProvider>
    {
        protected static ExcelColumn ResultExcelColumn;
    }

    internal class When_getting_column_by_column_number : ExcelColumnProviderSpecs
    {
        Because of = () => 
            ResultExcelColumn = sut.GetColumn(10, 5);

        It should_have_proper_width = () =>
            ResultExcelColumn.Width.ShouldBeGreaterThan(10);

        It should_have_proper_start_column_number = () =>
            ResultExcelColumn.ColumnNumberStart.ShouldEqual<uint>(5);

        It should_have_proper_end_column_number = () =>
            ResultExcelColumn.ColumnNumberEnd.ShouldEqual<uint>(5);
    }

    internal class When_getting_column_by_start_and_end_column_number : ExcelColumnProviderSpecs
    {
        Because of = () =>
            ResultExcelColumn = sut.GetColumn(20, 5, 7);

        It should_have_proper_width = () =>
            ResultExcelColumn.Width.ShouldBeGreaterThan(20);

        It should_have_proper_start_column_number = () =>
            ResultExcelColumn.ColumnNumberStart.ShouldEqual<uint>(5);

        It should_have_proper_end_column_number = () =>
            ResultExcelColumn.ColumnNumberEnd.ShouldEqual<uint>(7);
    }

    internal class When_getting_column_by_max_possible_characters : ExcelColumnProviderSpecs
    {
        Because of = () =>
            ResultExcelColumn = sut.GetColumn(int.MaxValue, 5);

        It should_have_proper_width = () =>
            ResultExcelColumn.Width.ShouldBeGreaterThan(int.MaxValue);

        It should_have_proper_start_column_number = () =>
            ResultExcelColumn.ColumnNumberStart.ShouldEqual<uint>(5);

        It should_have_proper_end_column_number = () =>
            ResultExcelColumn.ColumnNumberEnd.ShouldEqual<uint>(5);
    }

    internal class When_getting_column_by_min_possible_characters : ExcelColumnProviderSpecs
    {
        Because of = () =>
            ResultExcelColumn = sut.GetColumn(int.MinValue, 5);

        It should_have_proper_width = () =>
            ResultExcelColumn.Width.ShouldBeGreaterThan(int.MinValue);

        It should_have_proper_start_column_number = () =>
            ResultExcelColumn.ColumnNumberStart.ShouldEqual<uint>(5);

        It should_have_proper_end_column_number = () =>
            ResultExcelColumn.ColumnNumberEnd.ShouldEqual<uint>(5);
    }
}
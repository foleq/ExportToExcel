using System;
using System.Collections.Generic;
using Machine.Specifications;

namespace ExportToExcel.Tests.ExcelBuilderSpecs
{
    internal class When_manipulating_excel_after_creation_by_building_excel : ExcelBuilderSpecs
    {
        private static Exception _exception;

        private Because of = () =>
        {
            _exception = Catch.Exception((Action)(() => BuildExcel(ExpectedWorksheetDataList)));
            _exception = Catch.Exception(() => sut.AddRowToWorksheet("sheet_not_existing", new[] { "cell" }));
            _exception = Catch.Exception(() => sut.AddRowToWorksheet("sheet_1", new[] { "cell" }));
        };

        It should_be_an_InvalidOperationException = () =>
            _exception.ShouldBeOfExactType(typeof(InvalidOperationException));

        It should_have_proper_message = () =>
            _exception.Message.ShouldEqual(ExpectedExceptionMessage);
    }

    internal class When_manipulating_excel_after_creation_by_adding_row_to_existing_sheet : ExcelBuilderSpecs
    {
        private static Exception _exception;

        private Because of = () =>
        {
            _exception = Catch.Exception(() => sut.AddRowToWorksheet("sheet_1", new[] { "cell" }));
        };

        It should_be_an_InvalidOperationException = () =>
            _exception.ShouldBeOfExactType(typeof(InvalidOperationException));

        It should_have_proper_message = () =>
            _exception.Message.ShouldEqual(ExpectedExceptionMessage);
    }

    internal class When_manipulating_excel_after_creation_by_adding_row_to_not_existing_sheet : ExcelBuilderSpecs
    {
        private static Exception _exception;

        private Because of = () =>
        {
            _exception = Catch.Exception(() => sut.AddRowToWorksheet("sheet_not_existing", new[] { "cell" }));
        };

        It should_be_an_InvalidOperationException = () =>
            _exception.ShouldBeOfExactType(typeof(InvalidOperationException));

        It should_have_proper_message = () =>
            _exception.Message.ShouldEqual(ExpectedExceptionMessage);
    }

    internal class When_manipulating_excel_after_creation_by_again_getting_excel : ExcelBuilderSpecs
    {
        private static Exception _exception;

        private Because of = () =>
        {
            ResultDocument = BuildExcel(new List<ExpectedWorksheetData>());
        };

        It should_build_correctly_formatted_excel = () =>
            ResultDocument.ShouldNotBeNull();

        It should_have_proper_worksheets = () =>
            Should_have_proper_worksheets();

        It should_have_proper_data_for_worksheets = () =>
            Should_have_proper_data_for_worksheets();
    }
}
﻿using System;
using System.Collections.Generic;
using Machine.Specifications;

namespace ExportToExcel.Tests.ExcelBuilderSpecs
{
    internal abstract class When_manipulating_excel_after_creation : ExcelBuilderSpecs
    {
        Establish context = () =>
        {
            ExpectedWorksheetDataList = new List<ExpectedWorksheetData>()
            {
                new ExpectedWorksheetData()
                {
                    WorksheetIndex = 0,
                    WorksheetName = "sheet_1",
                    Data = new List<string[]>()
                    {
                        new[] { "row_1_cell_A", "row_1_cell_B" },
                        new[] { "row_2_cell_A", "row_2_cell_B" }
                    }
                }
            };
        };

        Because of = () =>
        {
            AddDataToExcel(ExpectedWorksheetDataList);
            ResultExcel = FinishAndGetResultExcel();
        };
    }

    internal class When_manipulating_excel_after_creation_by_adding_data_to_excel : When_manipulating_excel_after_creation
    {
        private static Exception _exception;

        Because of = () =>
        {
            _exception = Catch.Exception((Action)(() => AddDataToExcel(ExpectedWorksheetDataList) ));
        };

        It should_be_an_InvalidOperationException = () =>
            _exception.ShouldBeOfExactType(typeof(InvalidOperationException));

        It should_have_proper_message = () =>
            _exception.Message.ShouldEqual(ExpectedExceptionMessage);
    }

    internal class When_manipulating_excel_after_creation_by_adding_row_to_existing_sheet : When_manipulating_excel_after_creation
    {
        private static Exception _exception;

        Because of = () =>
        {
            _exception = Catch.Exception(() => sut.AddRowToWorksheet("sheet_1", new[] { "cell" }));
        };

        It should_be_an_InvalidOperationException = () =>
            _exception.ShouldBeOfExactType(typeof(InvalidOperationException));

        It should_have_proper_message = () =>
            _exception.Message.ShouldEqual(ExpectedExceptionMessage);
    }

    internal class When_manipulating_excel_after_creation_by_adding_row_to_not_existing_sheet : When_manipulating_excel_after_creation
    {
        private static Exception _exception;

        Because of = () =>
        {
            _exception = Catch.Exception(() => sut.AddRowToWorksheet("sheet_not_existing", new[] { "cell" }));
        };

        It should_be_an_InvalidOperationException = () =>
            _exception.ShouldBeOfExactType(typeof(InvalidOperationException));

        It should_have_proper_message = () =>
            _exception.Message.ShouldEqual(ExpectedExceptionMessage);
    }

    internal class When_manipulating_excel_after_creation_by_again_getting_excel : When_manipulating_excel_after_creation
    {
        Because of = () =>
        {
            ResultExcel = FinishAndGetResultExcel();
        };

        It should_build_correctly_formatted_excel = () =>
            ResultExcel.ShouldNotBeNull();

        It should_have_proper_worksheets = () =>
            Should_have_proper_worksheets();

        It should_have_proper_data_for_worksheets = () =>
            Should_have_proper_data_for_worksheets();

        It should_have_proper_stylesheet = () =>
            Should_have_proper_stylesheet();
    }
}
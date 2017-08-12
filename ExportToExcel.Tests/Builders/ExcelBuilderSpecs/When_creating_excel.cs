using System.Collections.Generic;
using Machine.Specifications;

namespace ExportToExcel.Tests.Builders.ExcelBuilderSpecs
{
    internal class When_creating_excel_with_one_worksheet : Builders.ExcelBuilderSpecs.ExcelBuilderSpecs
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

        It should_build_correctly_formatted_excel = () =>
            ResultExcel.ShouldNotBeNull();

        It should_have_proper_worksheets = () =>
            Should_have_proper_worksheets();

        It should_have_proper_data_for_worksheets = () =>
            Should_have_proper_data_for_worksheets();

        It should_have_proper_stylesheet = () =>
            Should_have_proper_stylesheet();
    }

    internal class When_creating_excel_with_multiple_worksheets : Builders.ExcelBuilderSpecs.ExcelBuilderSpecs
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
                },
                new ExpectedWorksheetData()
                {
                    WorksheetIndex = 1,
                    WorksheetName = "sheet_2",
                    Data = new List<string[]>()
                    {
                        new[] { "row_1_cell_A", "row_1_cell_B" },
                        new string[0],
                        new[] { "", "row_3_cell_A", "row_3_cell_B", "row_3_cell_C" },
                        new[] { "row_4_cell_A", "row_4_cell_B" }
                    }
                }
            };
        };

        Because of = () =>
        {
            AddDataToExcel(ExpectedWorksheetDataList);
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
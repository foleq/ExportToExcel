using System.Collections.Generic;
using ExportToExcel.Models;
using ExportToExcel.StylesheetProvider;
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
                    Data = new List<ExcelCell[]>()
                    {
                        new[]
                        {
                            new ExcelCell("row_1_cell_A", ExcelSheetStyleIndex.Bold),
                            new ExcelCell("2345678.99", ExcelSheetStyleIndex.Number),
                        },
                        new[] 
                        {
                            new ExcelCell("row_2_cell_A"),
                            new ExcelCell("row_2_cell_B")
                        }
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
                    Data = new List<ExcelCell[]>()
                    {
                        new[]
                        {
                            new ExcelCell("row_1_cell_A"),
                            new ExcelCell("row_1_cell_B"),
                        },
                        new[]
                        {
                            new ExcelCell("row_2_cell_A"),
                            new ExcelCell("row_2_cell_B")
                        }
                    }
                },
                new ExpectedWorksheetData()
                {
                    WorksheetIndex = 1,
                    WorksheetName = "sheet_2",
                    Data = new List<ExcelCell[]>()
                    {
                        new[]
                        {
                            new ExcelCell("row_1_cell_A"),
                            new ExcelCell("row_1_cell_B"),
                        },
                        new ExcelCell[0],
                        new[]
                        {
                            new ExcelCell(""),
                            new ExcelCell("row_3_cell_B"),
                            new ExcelCell("row_3_cell_C"),
                            new ExcelCell("row_3_cell_D"),
                        },
                        new[]
                        {
                            new ExcelCell("row_4_cell_A"),
                            new ExcelCell("row_4_cell_B")
                        }
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
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.Models;
using ExportToExcel.StylesheetProvider;
using Machine.Specifications;

namespace ExportToExcel.Tests.Builders.ExcelBuilderSpecs
{
    internal class When_creating_excel_with_one_worksheet : ExcelBuilderSpecs
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
                            new ExcelCell("2345678.99", ExcelSheetStyleIndex.Nformat4Decimal),
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

        It should_have_proper_columns_for_worksheets = () =>
            Should_have_proper_columns_for_worksheets(new Column[0]);

        It should_have_drawing_part_if_image_added = () =>
            Should_have_drawing_part_if_image_added();
    }

    internal class When_creating_excel_with_multiple_worksheets : ExcelBuilderSpecs
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

        It should_have_proper_columns_for_worksheets = () =>
            Should_have_proper_columns_for_worksheets(new Column[0]);

        It should_have_drawing_part_if_image_added = () =>
            Should_have_drawing_part_if_image_added();
    }

    internal class When_creating_excel_with_one_worksheet_with_auto_size_some_columns : ExcelBuilderSpecs
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
                            new ExcelCell(null, ExcelSheetStyleIndex.Default, true), 
                            new ExcelCell("some_very_long_column without_autosizing", ExcelSheetStyleIndex.Bold),
                            new ExcelCell("some_very_long_column some_very_long_column without_autosizing", ExcelSheetStyleIndex.Default, true),
                        },
                        new[]
                        {
                            new ExcelCell("", ExcelSheetStyleIndex.Default, true),
                            new ExcelCell("some_long_column with_autosizing", ExcelSheetStyleIndex.Default, true),
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

        It should_have_proper_columns_for_worksheets = () =>
            Should_have_proper_columns_for_worksheets(new[]
            {
                new Column() { Min = 3, Max = 3 },
                new Column() { Min = 2, Max = 2 },
            });

        It should_have_drawing_part_if_image_added = () =>
            Should_have_drawing_part_if_image_added();
    }

    internal class When_creating_excel_with_one_worksheet_and_image : ExcelBuilderSpecs
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
                        new ExcelCell[0],
                        new ExcelCell[0],
                        new ExcelCell[0],
                        new ExcelCell[0],
                        new[]
                        {
                            new ExcelCell("row_5_cell_B", ExcelSheetStyleIndex.Bold),
                            new ExcelCell("row_5_cell_C", ExcelSheetStyleIndex.Bold),
                        },
                        new[]
                        {
                            new ExcelCell("row_6_cell_B"),
                            new ExcelCell("row_6_cell_C")
                        }
                    },
                    Image = GetImage("Images/logo.png", ImagePartType.Png)
                },
                new ExpectedWorksheetData()
                {
                    WorksheetIndex = 1,
                    WorksheetName = "sheet_2",
                    Data = new List<ExcelCell[]>(),
                    Image = GetImage("wrongpath", ImagePartType.Png)
                },
                new ExpectedWorksheetData()
                {
                    WorksheetIndex = 2,
                    WorksheetName = "sheet_3",
                    Data = new List<ExcelCell[]>(),
                    Image = GetImage("Images/logo.jpg", ImagePartType.Jpeg, 2, 2)
                },
                new ExpectedWorksheetData()
                {
                    WorksheetIndex = 3,
                    WorksheetName = "sheet_4",
                    Data = new List<ExcelCell[]>()
                    {
                        new[]
                        {
                            null, 
                            new ExcelCell("row_1_cell_B"),
                            new ExcelCell("row_1_cell_C")
                        }
                    }
                },
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

        It should_have_proper_columns_for_worksheets = () =>
            Should_have_proper_columns_for_worksheets(new Column[0]);

        It should_have_drawing_part_if_image_added = () =>
            Should_have_drawing_part_if_image_added();
    }
}
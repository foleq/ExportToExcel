using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using developwithpassion.specifications.rhinomocks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Machine.Specifications;

namespace ExportToExcel.Tests
{
    //public class CleanExcelBuilderResult : ICleanupAfterEveryContextInAssembly
    //{
    //    public void AfterContextCleanup()
    //    {
    //        ExcelBuilderSpecs.ResultDocument.Close();
    //    }
    //}

    [Subject(typeof(ExcelBuilder))]
    internal abstract class ExcelBuilderSpecs : Observes<ExcelBuilder>
    {
        Because of = () =>
        {
            foreach (var worksheetData in ExpectedWorksheetDataList)
            {
                foreach (var dataRow in worksheetData.Data)
                {
                    sut.AddRowToWorksheet(worksheetData.WorksheetName, dataRow);
                }
            }
            var bytes = sut.FinishAndGetExcel();
            sut.Dispose();
            ResultDocument = GetSpreadsheetDocumentFrom(bytes);
        };

        private static bool ByteArrayToFile(string fileName, byte[] byteArray)
        {
            try
            {
                using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(byteArray, 0, byteArray.Length);
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static SpreadsheetDocument GetSpreadsheetDocumentFrom(byte[] bytes)
        {
            //const string resultPath = "ResultDocument.xlsx";
            const string resultPath = "C:\\Users\\piotr\\Desktop\\excel 2k10\\ResultDocument.xlsx";

            ByteArrayToFile(resultPath, bytes);
            using (var document = SpreadsheetDocument.Open(resultPath, true))
            {
                return (SpreadsheetDocument)document.Clone();
            }
        }

        private static List<string[]> GetData(int worksheetIndex)
        {
            var worksheetParts = ResultDocument.WorkbookPart.WorksheetParts.ToArray();
            var worksheet = worksheetParts[worksheetIndex].Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();

            var data = sheetData.Elements<Row>().Select(row =>
                {
                    return row.Elements<Cell>().Select(x => x.InnerText).ToArray();
                });
            return data.ToList();
        }


        protected static void Should_have_proper_worksheets()
        {
            ResultDocument.WorkbookPart.WorksheetParts.Count().ShouldEqual(ExpectedWorksheetDataList.Count);
            var sheets = ResultDocument.WorkbookPart.Workbook.Sheets.ChildElements.ToArray();
            sheets.Length.ShouldEqual(ExpectedWorksheetDataList.Count);

            foreach (var worksheetData in ExpectedWorksheetDataList)
            {
                var workshName = (sheets[worksheetData.WorksheetIndex] as Sheet).Name.Value;
                workshName.ShouldEqual(worksheetData.WorksheetName);
            }
        }

        protected static void Should_have_proper_data_for_worksheets()
        {
            foreach (var expectedWorksheetData in ExpectedWorksheetDataList)
            {
                var data = GetData(expectedWorksheetData.WorksheetIndex);
                var expectedData = expectedWorksheetData.Data;

                data.Count.ShouldEqual(expectedData.Count);

                for (var i = 0; i < data.Count; i++)
                {
                    data[i].Length.ShouldEqual(expectedData[i].Length);
                    for (var j = 0; j < data[i].Length; j++)
                    {
                        data[i][j].ShouldEqual(expectedData[i][j]);
                    }
                }
            }
        }

        protected static List<ExpectedWorksheetData> ExpectedWorksheetDataList;
        public static SpreadsheetDocument ResultDocument;

        protected class ExpectedWorksheetData
        {
            public int WorksheetIndex { get; set; }
            public string WorksheetName { get; set; }
            public List<string[]> Data { get; set; }
        }
    }

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
                    Data = new List<string[]>()
                    {
                        new[] { "row_1_cell_A", "row_1_cell_B" },
                        new[] { "row_2_cell_A", "row_2_cell_B" }
                    }
                }
            };
        };

        It should_build_correctly_formatted_excel = () =>
            ResultDocument.ShouldNotBeNull();

        It should_have_proper_worksheets = () =>
            Should_have_proper_worksheets();

        It should_have_proper_data_for_worksheets = () => 
            Should_have_proper_data_for_worksheets();
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

        It should_build_correctly_formatted_excel = () =>
            ResultDocument.ShouldNotBeNull();

        It should_have_proper_worksheets = () =>
            Should_have_proper_worksheets();

        It should_have_proper_data_for_worksheets = () =>
            Should_have_proper_data_for_worksheets();
    }
}
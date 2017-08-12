using System.Collections.Generic;
using System.IO;
using System.Linq;
using developwithpassion.specifications.rhinomocks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.StylesheetProvider;
using Machine.Specifications;
using Rhino.Mocks;

namespace ExportToExcel.Tests.ExcelBuilderSpecs
{
    //public class CleanExcelBuilderResult : ICleanupAfterEveryContextInAssembly
    //{
    //    public void AfterContextCleanup()
    //    {
    //        ExcelBuilderSpecs.ResultExcel.Close();
    //    }
    //}

    [Subject(typeof(ExcelBuilder))]
    internal abstract class ExcelBuilderSpecs : Observes<ExcelBuilder>
    {
        //const string resultPath = "ResultExcel.xlsx";
        private const string ResultPath = "C:\\Users\\piotr\\Desktop\\excel 2k10\\ResultExcel.xlsx";
        private static readonly Stylesheet Stylesheet = new Stylesheet();

        protected static string ExpectedExceptionMessage = "WorksheetPartBuilder has finished building and any adding is not allowed.";
        protected static List<ExpectedWorksheetData> ExpectedWorksheetDataList;
        public static SpreadsheetDocument ResultExcel;

        Establish context = () =>
        {
            var stylesheetProvider = depends.on<IExcelStylesheetProvider>();
            stylesheetProvider.Stub(x => x.GetStylesheet()).Return(Stylesheet);
        };

        protected static void AddDataToExcel(List<ExpectedWorksheetData> worksheetDataList)
        {
            foreach (var worksheetData in worksheetDataList)
            {
                foreach (var dataRow in worksheetData.Data)
                {
                    sut.AddRowToWorksheet(worksheetData.WorksheetName, dataRow);
                }
            }
        }

        protected static SpreadsheetDocument FinishAndGetResultExcel()
        {
            var bytes = sut.FinishAndGetExcel();
            sut.Dispose();
            return GetSpreadsheetDocumentFrom(bytes);
        }

        private static SpreadsheetDocument GetSpreadsheetDocumentFrom(byte[] bytes)
        {
            ByteArrayToFile(ResultPath, bytes);
            using (var document = SpreadsheetDocument.Open(ResultPath, true))
            {
                return (SpreadsheetDocument)document.Clone();
            }
        }

        private static void ByteArrayToFile(string fileName, byte[] byteArray)
        {
            using (var fileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                fileStream.Write(byteArray, 0, byteArray.Length);
            }
        }
        
        protected static void Should_have_proper_worksheets()
        {
            ResultExcel.WorkbookPart.WorksheetParts.Count().ShouldEqual(ExpectedWorksheetDataList.Count);
            var sheets = ResultExcel.WorkbookPart.Workbook.Sheets.ChildElements.ToArray();
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

        protected static void Should_have_proper_stylesheet()
        {
            var resutStylesheet = ResultExcel.WorkbookPart.WorkbookStylesPart.Stylesheet;
            resutStylesheet.OuterXml.ShouldEqual(Stylesheet.OuterXml);
        }

        private static List<string[]> GetData(int worksheetIndex)
        {
            var worksheetParts = ResultExcel.WorkbookPart.WorksheetParts.ToArray();
            var worksheet = worksheetParts[worksheetIndex].Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();

            var data = sheetData.Elements<Row>().Select(row =>
            {
                return row.Elements<Cell>().Select(x => x.InnerText).ToArray();
            });
            return data.ToList();
        }

        protected class ExpectedWorksheetData
        {
            public int WorksheetIndex { get; set; }
            public string WorksheetName { get; set; }
            public List<string[]> Data { get; set; }
        }
    }
}
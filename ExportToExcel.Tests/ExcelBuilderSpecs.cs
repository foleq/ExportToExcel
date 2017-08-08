using System.IO;
using developwithpassion.specifications.rhinomocks;
using DocumentFormat.OpenXml.Packaging;
using Machine.Specifications;

namespace ExportToExcel.Tests
{
    public class CleanExcelBuilderResult : ICleanupAfterEveryContextInAssembly
    {
        public void AfterContextCleanup()
        {
            ExcelBuilderSpecs.Result.Close();
        }
    }

    [Subject(typeof(ExcelBuilder))]
    internal abstract class ExcelBuilderSpecs : Observes<ExcelBuilder>
    {
        Because of = () =>
        {
            var excelBytes = sut.GetExcelBytes();
            sut.Dispose();
            Result = GetDocumentFrom(excelBytes);
        };

        private static SpreadsheetDocument GetDocumentFrom(byte[] bytes)
        {
            using (var stream = new MemoryStream(bytes))
            {
                return SpreadsheetDocument.Open(stream, false);
            }
        }

        public static SpreadsheetDocument Result;
    }

    internal class When_creating_excel : ExcelBuilderSpecs
    {
        It should_build_correctly_formatted_excel = () =>
        {
            Result.ShouldNotBeNull();
        };
    }
}
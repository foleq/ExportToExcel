using System.IO;
using ConsoleAppForExportToExcel.ExportToExcel;

namespace ConsoleAppForExportToExcel
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var myExcelBuilderFactor = new MyExcelBuilderFactory();
            var myExcelBuilder = myExcelBuilderFactor.CreateMyExcelBuilder();

            var byteArray = myExcelBuilder.BuildExcelFile(4, 250, 10000);
            const string fileName = "ExportedFile.xlsx";

            using (var fileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                fileStream.Write(byteArray, 0, byteArray.Length);
            }
        }
    }
}

using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ExportToExcel
{
    public interface IExcelBuilder
    {
        byte[] GetExcelBytes();
    }

    public class ExcelBuilder : IExcelBuilder, IDisposable
    {
        private readonly MemoryStream _memoryStream;
        private readonly SpreadsheetDocument _document;

        public ExcelBuilder()
        {
            _memoryStream = new MemoryStream();
            _document = SpreadsheetDocument.Create(_memoryStream, SpreadsheetDocumentType.Workbook);
            _document.AddWorkbookPart();
        }

        public byte[] GetExcelBytes()
        {
            _document.Close();
            return _memoryStream.ToArray();
        }

        public void Dispose()
        {
            _document?.Dispose();
            _memoryStream?.Dispose();
        }
    }
}

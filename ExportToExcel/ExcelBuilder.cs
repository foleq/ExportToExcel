using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.StylesheetProvider;

namespace ExportToExcel
{
    public interface IExcelBuilder : IDisposable
    {
        void AddRowToWorksheet(string worksheetName, string[] cellValues);
        byte[] FinishAndGetExcel();
    }

    public class ExcelBuilder : IExcelBuilder
    {
        private readonly IExcelStylesheetProvider _stylesheetProvider;
        private readonly MemoryStream _memoryStream;
        private readonly SpreadsheetDocument _document;
        private readonly Dictionary<string, WorksheetPartBuilder> _worksheetPartBuilders;
        private bool _buildingIsFinished;

        public ExcelBuilder(IExcelStylesheetProvider stylesheetProvider)
        {
            _stylesheetProvider = stylesheetProvider;
            _worksheetPartBuilders = new Dictionary<string, WorksheetPartBuilder>();
            _buildingIsFinished = false;

            _memoryStream = new MemoryStream();
            _document = SpreadsheetDocument.Create(_memoryStream, SpreadsheetDocumentType.Workbook);
            _document.AddWorkbookPart();
        }

        public void AddRowToWorksheet(string worksheetName, string[] cellValues)
        {
            ThrowExceptionIfBuildingIsFinished();

            if (_worksheetPartBuilders.ContainsKey(worksheetName) == false)
            {
                var worksheetBuilder = new WorksheetPartBuilder(_document.WorkbookPart.AddNewPart<WorksheetPart>());
                _worksheetPartBuilders.Add(worksheetName, worksheetBuilder);    
            }
            _worksheetPartBuilders[worksheetName].AddRow(cellValues);
        }

        private void ThrowExceptionIfBuildingIsFinished()
        {
            if (_buildingIsFinished)
            {
                throw new InvalidOperationException("WorksheetPartBuilder has finished building and any adding is not allowed.");
            }
        }

        public byte[] FinishAndGetExcel()
        {
            if (_buildingIsFinished == false)
            {
                FinishBuilding();
            }
            return _memoryStream.ToArray();
        }

        private void FinishBuilding()
        {
            AddStylesheet();
            AddSheets();
            _document.Close();
            _buildingIsFinished = true;
        }

        private void AddStylesheet()
        {
            var stylesPart = _document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            using (var writer = OpenXmlWriter.Create(stylesPart))
            {
                writer.WriteElement(_stylesheetProvider.GetStylesheet());
                writer.Close();
            }
        }

        private void AddSheets()
        {
            using (var writer = OpenXmlWriter.Create(_document.WorkbookPart))
            {
                writer.WriteStartElement(new Workbook());
                writer.WriteStartElement(new Sheets());

                uint sheetId = 1;
                foreach (var worksheetPartBuilder in _worksheetPartBuilders)
                {
                    var worksheetPart = worksheetPartBuilder.Value.FinishAndGetResult();
                    writer.WriteElement(new Sheet()
                    {
                        Name = worksheetPartBuilder.Key,
                        SheetId = sheetId++,
                        Id = _document.WorkbookPart.GetIdOfPart(worksheetPart)
                    });
                }

                writer.WriteEndElement(); // end Sheets
                writer.WriteEndElement(); // end Workbook
                writer.Close();
            }
        }

        public void Dispose()
        {
            _document?.Dispose();
            _memoryStream?.Dispose();
        }
    }
}
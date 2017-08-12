using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.Factories;
using ExportToExcel.Models;
using ExportToExcel.StylesheetProvider;

namespace ExportToExcel.Builders
{
    public interface IExcelBuilder : IDisposable
    {
        void AddRowToWorksheet(string worksheetName, ExcelCell[] cells);
        byte[] FinishAndGetExcel();
    }

    public class ExcelBuilder : IExcelBuilder
    {
        private readonly IExcelStylesheetProvider _stylesheetProvider;
        private readonly IExcelCellFactory _excelCellFactory;
        private readonly MemoryStream _memoryStream;
        private readonly SpreadsheetDocument _document;
        private readonly Dictionary<string, ExcelWorksheetPartBuilder> _worksheetPartBuilders;
        private bool _buildingIsFinished;

        public ExcelBuilder(IExcelStylesheetProvider stylesheetProvider,
            IExcelCellFactory excelCellFactory)
        {
            _stylesheetProvider = stylesheetProvider;
            _excelCellFactory = excelCellFactory;
            _worksheetPartBuilders = new Dictionary<string, ExcelWorksheetPartBuilder>();
            _buildingIsFinished = false;

            _memoryStream = new MemoryStream();
            _document = SpreadsheetDocument.Create(_memoryStream, SpreadsheetDocumentType.Workbook);
            _document.AddWorkbookPart();
        }

        public void AddRowToWorksheet(string worksheetName, ExcelCell[] cells)
        {
            ThrowExceptionIfBuildingIsFinished();

            if (_worksheetPartBuilders.ContainsKey(worksheetName) == false)
            {
                var worksheetBuilder = new ExcelWorksheetPartBuilder(_document.WorkbookPart.AddNewPart<WorksheetPart>(), _excelCellFactory);
                _worksheetPartBuilders.Add(worksheetName, worksheetBuilder);
            }
            _worksheetPartBuilders[worksheetName].AddRow(cells);
        }

        private void ThrowExceptionIfBuildingIsFinished()
        {
            if (_buildingIsFinished)
            {
                throw new InvalidOperationException("ExcelWorksheetPartBuilder has finished building and any adding is not allowed.");
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
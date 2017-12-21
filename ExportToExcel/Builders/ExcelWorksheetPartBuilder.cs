using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.Factories;
using ExportToExcel.Models;
using ExportToExcel.Providers;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;

namespace ExportToExcel.Builders
{
    internal class ExcelWorksheetPartBuilder : IDisposable
    {
        private readonly WorksheetPart _worksheetPart;
        private readonly IExcelCellFactory _excelCellFactory;
        private readonly IExcelCellNameProvider _excelCellNameProvider;
        private readonly OpenXmlWriter _writer;

        private readonly List<ExcelImage> _excelImages;
        private int currentRowNumber;
        private readonly IDictionary<string, Uri> _cellNamesWithUri;
        private bool _buildingIsFinished;

        public ExcelWorksheetPartBuilder(WorksheetPart worksheetPart,
            IExcelCellFactory excelCellFactory,
            IExcelCellNameProvider excelCellNameProvider,
            Columns columns = null)
        {
            _excelImages = new List<ExcelImage>();
            currentRowNumber = 0;
            _cellNamesWithUri = new Dictionary<string, Uri>();
            _buildingIsFinished = false;

            _worksheetPart = worksheetPart;
            _excelCellFactory = excelCellFactory;
            _excelCellNameProvider = excelCellNameProvider;
            _writer = OpenXmlWriter.Create(_worksheetPart);

            _writer.WriteStartElement(new Worksheet());
            if (columns != null)
            {
                _writer.WriteElement(columns);
            }
            _writer.WriteStartElement(new SheetData());
        }

        public WorksheetPart FinishAndGetResult()
        {
            if (_buildingIsFinished == false)
            {
                FinishBuilding();
            }
            return _worksheetPart;
        }

        private void FinishBuilding()
        {
            _writer.WriteEndElement(); // end SheetData
            if (_excelImages.Count > 0)
            {
                _writer.WriteElement(GetDrawing());
            }
            //TODO: add some unit tests for adding hyperlinks please :)
            AddHyperlinksToCells();
            _writer.WriteEndElement(); // end Worksheet
            _writer.Close();

            _buildingIsFinished = true;
        }

        private void AddHyperlinksToCells()
        {
            if (_cellNamesWithUri.Any() == false)
            {
                return;
            }

            var hyperlinks = new Hyperlinks();
            foreach (var cellNameWithUri in _cellNamesWithUri)
            {
                var hyperlinkId = "hyperlinkFor_" + cellNameWithUri.Key;
                hyperlinks.AppendChild(new Hyperlink()
                {
                    Reference = cellNameWithUri.Key,
                    Id = hyperlinkId
                });
                _worksheetPart.AddHyperlinkRelationship(cellNameWithUri.Value, true, hyperlinkId);
            }
            _writer.WriteElement(hyperlinks);
        }

        private Drawing GetDrawing()
        {
            var drawingsPartBuilder = new ExcelDrawingsPartBuilder();
            return drawingsPartBuilder.BuildDrawing(_worksheetPart, _excelImages);
        }

        public void AddRow(ExcelCell[] cells)
        {
            ThrowExceptionIfBuildingIsFinished();

            _writer.WriteStartElement(new Row());

            currentRowNumber++;
            var currentColumnNumber = 1;
            foreach (var cell in cells)
            {
                _writer.WriteElement(_excelCellFactory.GetCell(cell));
                TryAddUriForCell(cell, currentColumnNumber);
                currentColumnNumber++;
            }
            _writer.WriteEndElement(); // end Row 
        }

        private void TryAddUriForCell(ExcelCell cell, int currentColumnNumber)
        {
            if (cell == null || cell.Uri == null)
            {
                return;
            }
            var cellName = _excelCellNameProvider.GetCellName(currentColumnNumber, currentRowNumber);
            _cellNamesWithUri.Add(cellName, cell.Uri);
        }

        private void ThrowExceptionIfBuildingIsFinished()
        {
            if (_buildingIsFinished)
            {
                throw new InvalidOperationException("ExcelWorksheetPartBuilder has finished building and any adding is not allowed.");
            }
        }

        public void AddExcelImage(ExcelImage excelImage)
        {
            if (excelImage?.ImageBytes != null)
            {
                _excelImages.Add(excelImage);
            }
        }

        public void Dispose()
        {
            _writer.Dispose();
        }
    }
}
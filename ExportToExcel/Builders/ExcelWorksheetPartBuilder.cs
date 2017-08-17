using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.Factories;
using ExportToExcel.Models;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;

namespace ExportToExcel.Builders
{
    internal class ExcelWorksheetPartBuilder : IDisposable
    {
        private readonly WorksheetPart _worksheetPart;
        private readonly IExcelCellFactory _excelCellFactory;
        private readonly OpenXmlWriter _writer;

        private readonly List<ExcelImage> _excelImages;
        private bool _buildingIsFinished;

        public ExcelWorksheetPartBuilder(WorksheetPart worksheetPart,
            IExcelCellFactory excelCellFactory,
            Columns columns = null)
        {
            _excelImages = new List<ExcelImage>();
            _buildingIsFinished = false;

            _worksheetPart = worksheetPart;
            _excelCellFactory = excelCellFactory;
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
            _writer.WriteEndElement(); // end Worksheet
            _writer.Close();

            _buildingIsFinished = true;
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
            foreach (var cell in cells)
            {
                _writer.WriteElement(_excelCellFactory.GetCell(cell));
            }
            _writer.WriteEndElement(); // end Row 
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
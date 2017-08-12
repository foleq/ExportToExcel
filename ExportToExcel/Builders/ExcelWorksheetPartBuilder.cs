using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.Factories;
using ExportToExcel.Models;

namespace ExportToExcel.Builders
{
    internal class ExcelWorksheetPartBuilder : IDisposable
    {
        private readonly WorksheetPart _worksheetPart;
        private readonly IExcelCellFactory _excelCellFactory;
        private readonly OpenXmlWriter _writer;
        private bool _buildingIsFinished;

        public ExcelWorksheetPartBuilder(WorksheetPart worksheetPart,
            IExcelCellFactory excelCellFactory)
        {
            _buildingIsFinished = false;

            _worksheetPart = worksheetPart;
            _excelCellFactory = excelCellFactory;
            _writer = OpenXmlWriter.Create(_worksheetPart);

            _writer.WriteStartElement(new Worksheet());
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
            _writer.WriteEndElement(); // end Worksheet
            _writer.Close();

            _buildingIsFinished = true;
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

        public void Dispose()
        {
            _writer.Dispose();
        }
    }
}
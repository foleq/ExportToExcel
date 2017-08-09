using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportToExcel
{
    internal class WorksheetPartBuilder : IDisposable
    {
        private readonly WorksheetPart _worksheetPart;
        private readonly OpenXmlWriter _writer;
        private bool _buildingIsFinished;

        public WorksheetPartBuilder(WorksheetPart worksheetPart)
        {
            _buildingIsFinished = false;

            _worksheetPart = worksheetPart;
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

        public void AddRow(string[] cellValues)
        {
            ThrowExceptionIfBuildingIsFinished();

            _writer.WriteStartElement(new Row());
            foreach (var cellValue in cellValues)
            {
                _writer.WriteStartElement(new Cell()
                {
                    DataType = CellValues.String,
                });
                _writer.WriteElement(new CellValue(cellValue));
                _writer.WriteEndElement(); // end Cell
            }
            _writer.WriteEndElement(); // end Row 
        }

        private void ThrowExceptionIfBuildingIsFinished()
        {
            if (_buildingIsFinished)
            {
                throw new InvalidOperationException("WorksheetPartBuilder has finished building and any adding is not allowed.");
            }
        }

        public void Dispose()
        {
            _writer.Dispose();
        }
    }
}
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

        public WorksheetPartBuilder(WorksheetPart worksheetPart)
        {
            _worksheetPart = worksheetPart;
            _writer = OpenXmlWriter.Create(_worksheetPart);

            _writer.WriteStartElement(new Worksheet());
            _writer.WriteStartElement(new SheetData());
        }

        public WorksheetPart FinishBuilding()
        {
            _writer.WriteEndElement(); // end SheetData
            _writer.WriteEndElement(); // end Worksheet
            _writer.Close();

            return _worksheetPart;
        }

        public void AddRow(string[] cellValues)
        {
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

        public void Dispose()
        {
            _writer.Dispose();
        }
    }
}
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

        private readonly Dictionary<int, int> _maxNumberOfCharactersInColumns;
        private readonly List<ExcelImage> _excelImages;
        private bool _buildingIsFinished;

        public ExcelWorksheetPartBuilder(WorksheetPart worksheetPart,
            IExcelCellFactory excelCellFactory)
        {
            _maxNumberOfCharactersInColumns = new Dictionary<int, int>();
            _excelImages = new List<ExcelImage>();
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
            _writer.WriteElement(GetColumns());
            if (_excelImages.Count > 0)
            {
                _writer.WriteElement(GetDrawing());
            }
            _writer.WriteEndElement(); // end Worksheet
            _writer.Close();

            _buildingIsFinished = true;
        }

        private Columns GetColumns()
        {
            const double maxWidthOfFont = 7;
            var columns = new Columns();
            foreach (var maxNumberOfCharactersInColumn in _maxNumberOfCharactersInColumns)
            {
                // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column(v=office.14).aspx
                // width = Truncate([{Nformat4Decimal of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
                var width = Math.Truncate((maxNumberOfCharactersInColumn.Value * maxWidthOfFont + 5) / maxWidthOfFont * 256) / 256;

                var col = new Column()
                {
                    BestFit = true,
                    Min = (uint)(maxNumberOfCharactersInColumn.Key + 1),
                    Max = (uint)(maxNumberOfCharactersInColumn.Key + 1),
                    CustomWidth = true,
                    Width = width
                };
                columns.AppendChild(col);
            }
            return columns;
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

            CalculateMaxNumberOfCharactersInColumns(cells);
        }

        private void ThrowExceptionIfBuildingIsFinished()
        {
            if (_buildingIsFinished)
            {
                throw new InvalidOperationException("ExcelWorksheetPartBuilder has finished building and any adding is not allowed.");
            }
        }

        private void CalculateMaxNumberOfCharactersInColumns(IReadOnlyList<ExcelCell> cells)
        {
            for (var columnIndex = 0; columnIndex < cells.Count; columnIndex++)
            {
                if (cells[columnIndex] == null
                    || cells[columnIndex].WithAutoSize == false 
                    || string.IsNullOrEmpty(cells[columnIndex].Value))
                {
                    continue;
                }
                var cellTextLength = cells[columnIndex].Value.Length;
                if (_maxNumberOfCharactersInColumns.ContainsKey(columnIndex))
                {
                    var current = _maxNumberOfCharactersInColumns[columnIndex];
                    if (cellTextLength > current)
                    {
                        _maxNumberOfCharactersInColumns[columnIndex] = cellTextLength;
                    }
                }
                else
                {
                    _maxNumberOfCharactersInColumns.Add(columnIndex, cellTextLength);
                }
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
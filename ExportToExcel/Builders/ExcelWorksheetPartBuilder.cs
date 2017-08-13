using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.Factories;
using ExportToExcel.Models;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using OpenXmlDrawing = DocumentFormat.OpenXml.Drawing;

namespace ExportToExcel.Builders
{
    internal class ExcelWorksheetPartBuilder : IDisposable
    {
        private readonly WorksheetPart _worksheetPart;
        private readonly IExcelCellFactory _excelCellFactory;
        private readonly OpenXmlWriter _writer;

        private readonly Dictionary<int, int> _maxNumberOfCharactersInColumns;
        private ExcelImage _excelImage;
        private bool _buildingIsFinished;

        public ExcelWorksheetPartBuilder(WorksheetPart worksheetPart,
            IExcelCellFactory excelCellFactory)
        {
            _maxNumberOfCharactersInColumns = new Dictionary<int, int>();
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
            if (_excelImage?.ImageBytes != null)
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
            var drawingsPart = _worksheetPart.AddNewPart<DrawingsPart>();
            var worksheetDrawing = new WorksheetDrawing();

            var imagePart = drawingsPart.AddImagePart(_excelImage.Type);

            using (var stream = new MemoryStream(_excelImage.ImageBytes))
            {
                imagePart.FeedData(stream);
            }
            long extentsCx, extentsCy;
            using (var stream = new MemoryStream(_excelImage.ImageBytes))
            {
                var bm = new Bitmap(stream);
                extentsCx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
                extentsCy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
                bm.Dispose();
            }

            const int colOffset = 0;
            const int rowOffset = 0;

            var nvps = worksheetDrawing.Descendants<NonVisualDrawingProperties>();
            var nvpId = nvps.Any() ?
                (UInt32Value)worksheetDrawing.Descendants<NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1 :
                1U;

            var oneCellAnchor = new OneCellAnchor(
                new Xdr.FromMarker
                {
                    ColumnId = new ColumnId((_excelImage.ColNumber - 1).ToString()),
                    RowId = new RowId((_excelImage.RowNumber - 1).ToString()),
                    ColumnOffset = new ColumnOffset(colOffset.ToString()),
                    RowOffset = new RowOffset(rowOffset.ToString())
                },
                new Extent { Cx = extentsCx, Cy = extentsCy },
                new Xdr.Picture(
                    new NonVisualPictureProperties(
                        new NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId },
                        new NonVisualPictureDrawingProperties(new OpenXmlDrawing.PictureLocks { NoChangeAspect = true })
                    ),
                    new BlipFill(
                        new OpenXmlDrawing.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = OpenXmlDrawing.BlipCompressionValues.Print },
                        new OpenXmlDrawing.Stretch(new OpenXmlDrawing.FillRectangle())
                    ),
                    new ShapeProperties(
                        new OpenXmlDrawing.Transform2D(
                            new OpenXmlDrawing.Offset { X = 0, Y = 0 },
                            new OpenXmlDrawing.Extents { Cx = extentsCx, Cy = extentsCy }
                        ),
                        new OpenXmlDrawing.PresetGeometry { Preset = OpenXmlDrawing.ShapeTypeValues.Rectangle }
                    )
                ),
                new ClientData()
            );
            using (var drawingsPartWriter = OpenXmlWriter.Create(drawingsPart))
            {
                drawingsPartWriter.WriteStartElement(worksheetDrawing);
                drawingsPartWriter.WriteElement(oneCellAnchor);
                drawingsPartWriter.WriteEndElement();
                drawingsPartWriter.Close();
            }
            return new Drawing
            {
                Id = _worksheetPart.GetIdOfPart(drawingsPart)
            };
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

        public void SetExcelImage(ExcelImage excelImage)
        {
            _excelImage = excelImage;
        }

        public void Dispose()
        {
            _writer.Dispose();
        }
    }
}
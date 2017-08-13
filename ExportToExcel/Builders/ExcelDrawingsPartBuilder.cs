using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using ExportToExcel.Models;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;

namespace ExportToExcel.Builders
{
    internal class ExcelDrawingsPartBuilder
    {
        public Drawing BuildDrawing(WorksheetPart worksheetPart, List<ExcelImage> excelImages)
        {
            var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

            var worksheetDrawing = new WorksheetDrawing();
            var oneCellAnchors = new List<OneCellAnchor>();

            foreach (var excelImage in excelImages)
            {
                var imagePart = drawingsPart.AddImagePart(excelImage.Type);

                using (var stream = new MemoryStream(excelImage.ImageBytes))
                {
                    imagePart.FeedData(stream);
                }
                long extentsCx, extentsCy;
                using (var stream = new MemoryStream(excelImage.ImageBytes))
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
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker
                    {
                        ColumnId = new ColumnId((excelImage.ColNumber - 1).ToString()),
                        RowId = new RowId((excelImage.RowNumber - 1).ToString()),
                        ColumnOffset = new ColumnOffset(colOffset.ToString()),
                        RowOffset = new RowOffset(rowOffset.ToString())
                    },
                    new Extent { Cx = extentsCx, Cy = extentsCy },
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture(
                        new NonVisualPictureProperties(
                            new NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId },
                            new NonVisualPictureDrawingProperties(new DocumentFormat.OpenXml.Drawing.PictureLocks { NoChangeAspect = true })
                        ),
                        new BlipFill(
                            new DocumentFormat.OpenXml.Drawing.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print },
                            new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())
                        ),
                        new ShapeProperties(
                            new DocumentFormat.OpenXml.Drawing.Transform2D(
                                new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
                                new DocumentFormat.OpenXml.Drawing.Extents { Cx = extentsCx, Cy = extentsCy }
                            ),
                            new DocumentFormat.OpenXml.Drawing.PresetGeometry { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }
                        )
                    ),
                    new ClientData()
                );
                oneCellAnchors.Add(oneCellAnchor);
            }

            using (var drawingsPartWriter = OpenXmlWriter.Create(drawingsPart))
            {
                drawingsPartWriter.WriteStartElement(worksheetDrawing);
                foreach (var oneCellAnchor in oneCellAnchors)
                {
                    drawingsPartWriter.WriteElement(oneCellAnchor);

                }
                drawingsPartWriter.WriteEndElement();
                drawingsPartWriter.Close();
            }

            return new Drawing
            {
                Id = worksheetPart.GetIdOfPart(drawingsPart)
            };
        }
    }
}
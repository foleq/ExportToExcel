using DocumentFormat.OpenXml.Packaging;

namespace ExportToExcel.Models
{
    public class ExcelImage
    {
        public byte[] ImageBytes { get; }
        public ImagePartType Type { get; }
        public int ColNumber { get; }
        public int RowNumber { get; }

        public ExcelImage(byte[] imageBytes, ImagePartType type = ImagePartType.Png, int colNumber = 1, int rowNumber = 1)
        {
            ImageBytes = imageBytes;
            Type = type;
            ColNumber = colNumber;
            RowNumber = rowNumber;
        }
    }
}
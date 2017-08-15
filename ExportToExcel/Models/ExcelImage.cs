using DocumentFormat.OpenXml.Packaging;

namespace ExportToExcel.Models
{
    public enum ExcelImageType
    {
        Png,
        Jpeg
    }

    public class ExcelImage
    {
        public byte[] ImageBytes { get; }
        internal ImagePartType Type { get; }
        public int ColNumber { get; }
        public int RowNumber { get; }

        public ExcelImage(byte[] imageBytes, ExcelImageType type = ExcelImageType.Png, int colNumber = 1, int rowNumber = 1)
        {
            ImageBytes = imageBytes;
            ColNumber = colNumber;
            RowNumber = rowNumber;
            if (type == ExcelImageType.Png)
            {
                Type = ImagePartType.Png;
            }
            else if (type == ExcelImageType.Jpeg)
            {
                Type = ImagePartType.Jpeg;
            }
        }
    }
}
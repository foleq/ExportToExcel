using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportToExcel.StylesheetProvider
{
    public enum ExcelSheetStyleIndex
    {
        Default,
        Bold,
        Number,
    }

    public interface IExcelStylesheetProvider
    {
        uint GetSheetStyleIndex(ExcelSheetStyleIndex excelSheetStyleIndex);
        Stylesheet GetStylesheet();
    }

    public class ExcelStylesheetProvider : IExcelStylesheetProvider
    {
        private readonly IExcelStylesheetNumberingFormatProvider _numberingFormatProvider;
        private readonly IExcelStylesheetFontProvider _fontProvider;
        private readonly IExcelStylesheetFillProvider _fillProvider;
        private readonly IExcelStylesheetBorderProvider _borderProvider;
        private readonly Stylesheet _stylesheet;
        private static IDictionary<ExcelSheetStyleIndex, uint> _indexes;

        public ExcelStylesheetProvider(IExcelStylesheetNumberingFormatProvider numberingFormatProvider,
            IExcelStylesheetFontProvider fontProvider,
            IExcelStylesheetFillProvider fillProvider,
            IExcelStylesheetBorderProvider borderProvider)
        {
            _numberingFormatProvider = numberingFormatProvider;
            _fontProvider = fontProvider;
            _fillProvider = fillProvider;
            _borderProvider = borderProvider;

            _indexes = new Dictionary<ExcelSheetStyleIndex, uint>();
            _stylesheet = CreateStylesheet();
        }

        private Stylesheet CreateStylesheet()
        {
            var stylesheet = new Stylesheet();
            // Order is important!
            stylesheet.AppendChild(_numberingFormatProvider.GetNumberingFormats());
            stylesheet.AppendChild(_fontProvider.GetFonts());
            stylesheet.AppendChild(_fillProvider.GetFills());
            stylesheet.AppendChild(_borderProvider.GetBorders());
            stylesheet.AppendChild(CreateCellFormats());
            return stylesheet;
        }

        private CellFormats CreateCellFormats()
        {
            var cellformats = new CellFormats();

            var defaultCellFromat = new CellFormat()
            {
                FontId = _fontProvider.GetFontIndex(ExcelStylesheetFontIndex.Default),
                FillId = _fillProvider.GetFillIndex(ExcelStylesheetFillIndex.Default),
                BorderId = _borderProvider.GetBorderIndex(ExcelSheetBorderIndex.Default),
            };
            AppendWithIndexSave(cellformats, defaultCellFromat, ExcelSheetStyleIndex.Default);

            var boldCellFormat = new CellFormat()
            {
                FontId = _fontProvider.GetFontIndex(ExcelStylesheetFontIndex.Bold),
                ApplyFont = true
            };
            AppendWithIndexSave(cellformats, boldCellFormat, ExcelSheetStyleIndex.Bold);

            var numberCellFromat = new CellFormat()
            {
                NumberFormatId = _numberingFormatProvider.GetNumberFormatIndex(ExcelSheetNumberingFormatIndex.Nformat4Decimal),
                ApplyNumberFormat = true
            };
            AppendWithIndexSave(cellformats, numberCellFromat, ExcelSheetStyleIndex.Number);

            return cellformats;
        }

        private void AppendWithIndexSave(CellFormats formats, CellFormat child, ExcelSheetStyleIndex excelSheetIndex)
        {
            formats.AppendChild(child);
            _indexes.Add(excelSheetIndex, (uint)_indexes.Count);
        }

        public uint GetSheetStyleIndex(ExcelSheetStyleIndex index)
        {
            if (_indexes != null && _indexes.ContainsKey(index))
            {
                return _indexes[index];
            }
            return 0;
        }

        public Stylesheet GetStylesheet()
        {
            return _stylesheet;
        }
    }
}
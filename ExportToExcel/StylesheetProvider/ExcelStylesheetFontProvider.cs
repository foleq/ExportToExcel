using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportToExcel.StylesheetProvider
{
    public enum ExcelStylesheetFontIndex
    {
        Default,
        Bold,
    }

    public interface IExcelStylesheetFontProvider
    {
        Fonts GetFonts();
        uint GetFontIndex(ExcelStylesheetFontIndex index);
    }

    public class ExcelStylesheetFontProvider : IExcelStylesheetFontProvider
    {
        private readonly Fonts _fonts;
        private readonly IDictionary<ExcelStylesheetFontIndex, uint> _indexes;

        public ExcelStylesheetFontProvider()
        {
            _indexes = new Dictionary<ExcelStylesheetFontIndex, uint>();
            _fonts = CreateFonts();
        }

        private Fonts CreateFonts()
        {
            var defaultFont = new Font();
            var boldFont = new Font();
            boldFont.AppendChild(new Bold());
            var fonts = new Fonts();
            AppendWithIndexSave(fonts, defaultFont, ExcelStylesheetFontIndex.Default);
            AppendWithIndexSave(fonts, boldFont, ExcelStylesheetFontIndex.Bold);
            return fonts;
        }

        private void AppendWithIndexSave(Fonts parent, Font child, ExcelStylesheetFontIndex excelSheetIndex)
        {
            parent.AppendChild(child);
            _indexes.Add(excelSheetIndex, (uint)_indexes.Count);
        }

        public Fonts GetFonts()
        {
            return _fonts;
        }

        public uint GetFontIndex(ExcelStylesheetFontIndex index)
        {
            if (_indexes != null && _indexes.ContainsKey(index))
            {
                return _indexes[index];
            }
            return 0;
        }
    }
}
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportToExcel.StylesheetProvider
{
    public enum ExcelSheetNumberingFormatIndex
    {
        Nformat4Decimal,
    }

    public interface IExcelStylesheetNumberingFormatProvider
    {
        NumberingFormats GetNumberingFormats();
        uint GetNumberFormatIndex(ExcelSheetNumberingFormatIndex index);
    }

    public class ExcelStylesheetNumberingFormatProvider : IExcelStylesheetNumberingFormatProvider
    {
        private readonly NumberingFormats _numberingFormats;
        private readonly IDictionary<ExcelSheetNumberingFormatIndex, uint> _indexes;

        public ExcelStylesheetNumberingFormatProvider()
        {
            _indexes = new Dictionary<ExcelSheetNumberingFormatIndex, uint>();
            _numberingFormats = CreateNumberingFormats();
        }

        private NumberingFormats CreateNumberingFormats()
        {
            var nformat4Decimal = new NumberingFormat
            {
                NumberFormatId = 0,
                FormatCode = StringValue.FromString("#,##0.00")
            };
            var numberingFormats = new NumberingFormats();
            AppendWithIndexSave(numberingFormats, nformat4Decimal, ExcelSheetNumberingFormatIndex.Nformat4Decimal);
            return numberingFormats;
        }

        private void AppendWithIndexSave(NumberingFormats parent, NumberingFormat child, ExcelSheetNumberingFormatIndex excelSheetIndex)
        {
            parent.AppendChild(child);
            _indexes.Add(excelSheetIndex, (uint)_indexes.Count);
        }

        public NumberingFormats GetNumberingFormats()
        {
            return _numberingFormats;
        }

        public uint GetNumberFormatIndex(ExcelSheetNumberingFormatIndex index)
        {
            if (_indexes != null && _indexes.ContainsKey(index))
            {
                return _indexes[index];
            }
            return 0;
        }
    }
}
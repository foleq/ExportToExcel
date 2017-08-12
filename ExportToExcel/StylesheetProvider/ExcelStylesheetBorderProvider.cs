using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportToExcel.StylesheetProvider
{
    public enum ExcelSheetBorderIndex
    {
        Default,
    }

    public interface IExcelStylesheetBorderProvider
    {
        Borders GetBorders();
        uint GetBorderIndex(ExcelSheetBorderIndex index);
    }

    public class ExcelStylesheetBorderProvider : IExcelStylesheetBorderProvider
    {
        private readonly Borders _borders;
        private readonly IDictionary<ExcelSheetBorderIndex, uint> _indexes;

        public ExcelStylesheetBorderProvider()
        {
            _indexes = new Dictionary<ExcelSheetBorderIndex, uint>();
            _borders = CreateBorders();
        }

        private Borders CreateBorders()
        {
            var defaultBorder = new Border();
            var borders = new Borders();
            AppendWithIndexSave(borders, defaultBorder, ExcelSheetBorderIndex.Default);
            return borders;
        }

        private void AppendWithIndexSave(Borders parent, Border child, ExcelSheetBorderIndex excelSheetIndex)
        {
            parent.AppendChild(child);
            _indexes.Add(excelSheetIndex, (uint)_indexes.Count);
        }

        public Borders GetBorders()
        {
            return _borders;
        }

        public uint GetBorderIndex(ExcelSheetBorderIndex index)
        {
            if (_indexes != null && _indexes.ContainsKey(index))
            {
                return _indexes[index];
            }
            return 0;
        }
    }
}
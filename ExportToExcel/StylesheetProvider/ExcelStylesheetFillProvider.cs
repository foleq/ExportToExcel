using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportToExcel.StylesheetProvider
{
    public enum ExcelStylesheetFillIndex
    {
        Default,
    }

    public interface IExcelStylesheetFillProvider
    {
        Fills GetFills();
        uint GetFillIndex(ExcelStylesheetFillIndex index);
    }

    public class ExcelStylesheetFillProvider : IExcelStylesheetFillProvider
    {
        private readonly Fills _fills;
        private readonly IDictionary<ExcelStylesheetFillIndex, uint> _indexes;

        public ExcelStylesheetFillProvider()
        {
            _indexes = new Dictionary<ExcelStylesheetFillIndex, uint>();
            _fills = CreateFills();
        }

        private Fills CreateFills()
        {
            var defaultFill = new Fill();
            var fills = new Fills();
            AppendWithIndexSave(fills, defaultFill, ExcelStylesheetFillIndex.Default);
            return fills;
        }

        private void AppendWithIndexSave(Fills parent, Fill child, ExcelStylesheetFillIndex excelStylesheetIndex)
        {
            parent.AppendChild(child);
            _indexes.Add(excelStylesheetIndex, (uint)_indexes.Count);
        }

        public Fills GetFills()
        {
            return _fills;
        }

        public uint GetFillIndex(ExcelStylesheetFillIndex index)
        {
            if (_indexes != null && _indexes.ContainsKey(index))
            {
                return _indexes[index];
            }
            return 0;
        }
    }
}
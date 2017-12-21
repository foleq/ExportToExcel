using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.Factories;
using ExportToExcel.Models;
using ExportToExcel.Providers;
using ExportToExcel.StylesheetProvider;

namespace ExportToExcel.Builders
{
    public interface IExcelBuilder : IDisposable
    {
        void AddWorksheet(string worksheetName, ExcelColumn[] columns = null);
        void AddRowToWorksheet(string worksheetName, ExcelCell[] cells);
        void AddImage(string worksheetName, ExcelImage excelImage);
        byte[] FinishAndGetExcel();
    }

    public class ExcelBuilder : IExcelBuilder
    {
        private readonly IExcelStylesheetProvider _stylesheetProvider;
        private readonly IExcelCellFactory _excelCellFactory;
        private readonly IExcelCellNameProvider _excelCellNameProvider;
        private readonly MemoryStream _memoryStream;
        private readonly SpreadsheetDocument _document;
        private readonly Dictionary<string, ExcelWorksheetPartBuilder> _worksheetPartBuilders;
        private bool _buildingIsFinished;

        public ExcelBuilder(IExcelStylesheetProvider stylesheetProvider,
            IExcelCellFactory excelCellFactory,
            IExcelCellNameProvider excelCellNameProvider)
        {
            _stylesheetProvider = stylesheetProvider;
            _excelCellFactory = excelCellFactory;
            _excelCellNameProvider = excelCellNameProvider;
            _worksheetPartBuilders = new Dictionary<string, ExcelWorksheetPartBuilder>();
            _buildingIsFinished = false;

            _memoryStream = new MemoryStream();
            _document = SpreadsheetDocument.Create(_memoryStream, SpreadsheetDocumentType.Workbook);
            _document.AddWorkbookPart();
        }

        public void AddWorksheet(string worksheetName, ExcelColumn[] columns = null)
        {
            if (_worksheetPartBuilders.ContainsKey(worksheetName))
            {
                throw new InvalidOperationException($"Worksheet with name '{worksheetName}' already exist in ExcelWorksheetPartBuilder.");
            }
            CreateWorksheetPartBuilderIfNotExist(worksheetName, columns);
        }

        public void AddRowToWorksheet(string worksheetName, ExcelCell[] cells)
        {
            ThrowExceptionIfBuildingIsFinished();

            CreateWorksheetPartBuilderIfNotExist(worksheetName);
            _worksheetPartBuilders[worksheetName].AddRow(cells);
        }

        public void AddImage(string worksheetName, ExcelImage excelImage)
        {
            CreateWorksheetPartBuilderIfNotExist(worksheetName);
            _worksheetPartBuilders[worksheetName].AddExcelImage(excelImage);
        }

        private void ThrowExceptionIfBuildingIsFinished()
        {
            if (_buildingIsFinished)
            {
                throw new InvalidOperationException("ExcelWorksheetPartBuilder has finished building and any adding is not allowed.");
            }
        }

        private void CreateWorksheetPartBuilderIfNotExist(string worksheetName, ExcelColumn[] columns = null)
        {
            if (_worksheetPartBuilders.ContainsKey(worksheetName))
            {
                return;
            }
            var worksheetBuilder = new ExcelWorksheetPartBuilder(_document.WorkbookPart.AddNewPart<WorksheetPart>(), 
                _excelCellFactory, _excelCellNameProvider, GetColumns(columns));
            _worksheetPartBuilders.Add(worksheetName, worksheetBuilder);
        }

        private static Columns GetColumns(ExcelColumn[] columns)
        {
            if (columns == null)
            {
                return null;
            }
            var cols = columns.Select(col => new Column()
            {
                BestFit = true,
                Min = col.ColumnNumberStart,
                Max = col.ColumnNumberEnd,
                CustomWidth = true,
                Width = col.Width
            });
            return new Columns(cols);
        }

        public byte[] FinishAndGetExcel()
        {
            if (_buildingIsFinished == false)
            {
                FinishBuilding();
            }
            return _memoryStream.ToArray();
        }

        private void FinishBuilding()
        {
            AddStylesheet();
            AddSheets();
            _document.Close();
            _buildingIsFinished = true;
        }

        private void AddStylesheet()
        {
            var stylesPart = _document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            using (var writer = OpenXmlWriter.Create(stylesPart))
            {
                writer.WriteElement(_stylesheetProvider.GetStylesheet());
                writer.Close();
            }
        }

        private void AddSheets()
        {
            using (var writer = OpenXmlWriter.Create(_document.WorkbookPart))
            {
                writer.WriteStartElement(new Workbook());
                writer.WriteStartElement(new Sheets());

                uint sheetId = 1;
                foreach (var worksheetPartBuilder in _worksheetPartBuilders)
                {
                    var worksheetPart = worksheetPartBuilder.Value.FinishAndGetResult();
                    writer.WriteElement(new Sheet()
                    {
                        Name = worksheetPartBuilder.Key,
                        SheetId = sheetId++,
                        Id = _document.WorkbookPart.GetIdOfPart(worksheetPart)
                    });
                }

                writer.WriteEndElement(); // end Sheets
                writer.WriteEndElement(); // end Workbook
                writer.Close();
            }
        }

        public void Dispose()
        {
            _document?.Dispose();
            _memoryStream?.Dispose();
        }
    }
}
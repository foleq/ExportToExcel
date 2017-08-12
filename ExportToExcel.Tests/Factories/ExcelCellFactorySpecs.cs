using developwithpassion.specifications.rhinomocks;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcel.Factories;
using ExportToExcel.Models;
using ExportToExcel.StylesheetProvider;
using Machine.Specifications;
using Rhino.Mocks;

namespace ExportToExcel.Tests.Factories
{
    [Subject(typeof(ExcelCellFactory))]
    internal abstract class ExcelCellFactorySpecs : Observes<ExcelCellFactory>
    {
        Establish context = () =>
        {
            var stylesheetProvider = depends.on<IExcelStylesheetProvider>();
            stylesheetProvider.Stub(x => x.GetSheetStyleIndex(ExcelSheetStyleIndex.Default))
                .Return(100);
            stylesheetProvider.Stub(x => x.GetSheetStyleIndex(ExcelSheetStyleIndex.Bold))
                .Return(101);
            stylesheetProvider.Stub(x => x.GetSheetStyleIndex(ExcelSheetStyleIndex.Number))
                .Return(102);
        };

        protected static Cell ResultCell;
    }

    internal class When_getting_cell_from_null : ExcelCellFactorySpecs
    {
        Because of = () =>
            ResultCell = sut.GetCell(null);

        It should_NOT_have_cell_value = () =>
            ResultCell.CellValue.ShouldBeNull();

        It should_NOT_have_data_type = () =>
            ResultCell.DataType.ShouldBeNull();

        It should_NOT_have_style_index = () =>
            ResultCell.StyleIndex.ShouldBeNull();
    }

    internal class When_getting_cell_without_value : ExcelCellFactorySpecs
    {
        Because of = () =>
            ResultCell = sut.GetCell(new ExcelCell(null, ExcelSheetStyleIndex.Default));

        It should_NOT_have_cell_value = () =>
            ResultCell.CellValue.ShouldBeNull();

        It should_NOT_have_data_type = () =>
            ResultCell.DataType.ShouldBeNull();

        It should_NOT_have_style_index = () =>
            ResultCell.StyleIndex.ShouldBeNull();
    }

    internal class When_getting_cell_with_string_value_and_default_style : ExcelCellFactorySpecs
    {
        Because of = () =>
            ResultCell = sut.GetCell(new ExcelCell("test_value", ExcelSheetStyleIndex.Default));

        It should_have_cell_value = () =>
            ResultCell.CellValue.InnerText.ShouldEqual("test_value");

        It should_have_data_type = () =>
            ResultCell.DataType.Value.ShouldEqual(CellValues.String);

        It should_have_style_index = () =>
            ResultCell.StyleIndex.Value.ShouldEqual<uint>(100);
    }

    internal class When_getting_cell_with_string_value_and_bold_style : ExcelCellFactorySpecs
    {
        Because of = () =>
            ResultCell = sut.GetCell(new ExcelCell("bold_value", ExcelSheetStyleIndex.Bold));

        It should_have_cell_value = () =>
            ResultCell.CellValue.InnerText.ShouldEqual("bold_value");

        It should_have_data_type = () =>
            ResultCell.DataType.Value.ShouldEqual(CellValues.String);

        It should_have_style_index = () =>
            ResultCell.StyleIndex.Value.ShouldEqual<uint>(101);
    }

    internal class When_getting_cell_with_numeric_value : ExcelCellFactorySpecs
    {
        Because of = () =>
            ResultCell = sut.GetCell(new ExcelCell((23456.99).ToString(), ExcelSheetStyleIndex.Number));

        It should_have_cell_value = () =>
            ResultCell.CellValue.InnerText.ShouldEqual("23456.99");

        It should_have_data_type = () =>
            ResultCell.DataType.Value.ShouldEqual(CellValues.Number);

        It should_have_style_index = () =>
            ResultCell.StyleIndex.Value.ShouldEqual<uint>(102);
    }
}
using System;
using System.Collections.Generic;
using developwithpassion.specifications.rhinomocks;
using ExportToExcel.Providers;
using Machine.Specifications;

namespace ExportToExcel.Tests.Providers
{
    [Subject(typeof(ExcelCellNameProvider))]
    internal abstract class ExcelCellNameProviderSpecs : Observes<ExcelCellNameProvider>
    {
    }

    internal class when_getting_cell_name_for_column_number_equal_0 : ExcelCellNameProviderSpecs
    {
        Because of = () =>
            exception = Catch.Exception(() =>
            {
                sut.GetCellName(0, 1);
            });

        It should_be_an_ArgumentException = () =>
            exception.ShouldBeOfExactType(typeof(ArgumentException));

        It should_have_proper_exception_message = () =>
            exception.Message.ShouldEqual("Column and row number should be greater than 0.");

        private static Exception exception;
    }

    internal class when_getting_column_name_for_column_number_between_1_and_26 : ExcelCellNameProviderSpecs
    {
        Because of = () =>
        {
            _columnNames = new List<string>()
            {
                sut.GetCellName(1, 11),
                sut.GetCellName(5, 22),
                sut.GetCellName(8, 33),
                sut.GetCellName(26, 44),
            };
        };

        It should_contain_proper_column_names_in_order = () =>
        {
            _columnNames[0].ShouldEqual("A11");
            _columnNames[1].ShouldEqual("E22");
            _columnNames[2].ShouldEqual("H33");
            _columnNames[3].ShouldEqual("Z44");
        };

        private static List<string> _columnNames;
    }

    internal class when_getting_column_name_for_column_number_greater_than_26 : ExcelCellNameProviderSpecs
    {
        Because of = () =>
        {
            _columnNames = new List<string>()
            {
                sut.GetCellName(26 + 1, 11),
                sut.GetCellName(26*2 + 5, 13),
                sut.GetCellName(26*2 + 26, 15),
                sut.GetCellName(26*26 + 26 + 2, 17),
            };
        };

        It should_contain_proper_column_names_in_order = () =>
        {
            _columnNames[0].ShouldEqual("AA11");
            _columnNames[1].ShouldEqual("BE13");
            _columnNames[2].ShouldEqual("BZ15");
            _columnNames[3].ShouldEqual("AAB17");
        };

        private static List<string> _columnNames;
    }
}
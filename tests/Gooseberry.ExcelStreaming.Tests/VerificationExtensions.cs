using System.Collections.Generic;
using System.Linq;
using Gooseberry.ExcelStreaming.Tests.Excel;
using FluentAssertions;
using FluentAssertions.Execution;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Tests
{
    public static class VerificationExtensions
    {
        public static void ShouldBeEquivalentTo(this IReadOnlyCollection<Sheet> actualSheets, params Sheet[] expectedSheets)
        {
            using var scope = new AssertionScope();

            actualSheets.Should().HaveCount(expectedSheets.Length);
            foreach (var (actual, expected) in actualSheets.Zip(expectedSheets))
            {
                actual.Name.Should().BeEquivalentTo(expected.Name);

                actual.Rows.Should().HaveCount(expected.Rows.Count);
                foreach (var (actualRow, expectedRow) in actual.Rows.Zip(expected.Rows))
                {
                    actualRow.Height.Should().Be(expectedRow.Height);
                    actualRow.Cells.Should().BeEquivalentTo(expectedRow.Cells);
                }

                actual.Columns.Should().HaveCount(expected.Columns.Count);
                foreach (var (actualColumn, expectedColumn) in actual.Columns.Zip(expected.Columns))
                    actualColumn.Should().BeEquivalentTo(expectedColumn);
            }
        }

        public static void ShouldBeEquivalentTo(this IReadOnlyCollection<Style> actualStyles, params Style[] expectedStyles)
        {
            using var scope = new AssertionScope();

            actualStyles.Should().HaveCount(expectedStyles.Length);
            foreach (var (actual, expected) in actualStyles.Zip(expectedStyles))
            {
                actual.Format.Should().BeEquivalentTo(expected.Format);
                actual.Fill.Should().BeEquivalentTo(expected.Fill);
                actual.Font.Should().BeEquivalentTo(expected.Font);
                actual.Borders.Should().BeEquivalentTo(expected.Borders);
                actual.Alignment.Should().BeEquivalentTo(expected.Alignment);
            }
        }
    }
}

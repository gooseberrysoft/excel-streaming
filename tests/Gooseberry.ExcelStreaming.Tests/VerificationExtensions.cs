using System.Drawing;
using FluentAssertions;
using FluentAssertions.Execution;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Tests;

internal static class VerificationExtensions
{
    static VerificationExtensions()
    {
        AssertionOptions
            .AssertEquivalencyUsing(options => options
                .Using<Color>(c => c.Subject.ToArgb().Should().Be(c.Expectation.ToArgb())).WhenTypeIs<Color>());
    }

    public static void ShouldBeEquivalentTo(this IReadOnlyCollection<Excel.Sheet> actualSheets, params Excel.Sheet[] expectedSheets)
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
                actualRow.OutlineLevel.Should().Be(expectedRow.OutlineLevel);
                actualRow.IsCollapsed.Should().Be(expectedRow.IsCollapsed);
                actualRow.IsHidden.Should().Be(expectedRow.IsHidden);
                actualRow.Cells.Should().BeEquivalentTo(expectedRow.Cells);
            }

            actual.Columns.Should().HaveCount(expected.Columns.Count);
            foreach (var (actualColumn, expectedColumn) in actual.Columns.Zip(expected.Columns))
                actualColumn.Should().BeEquivalentTo(expectedColumn);

            actual.Merges.Should().HaveCount(expected.Merges.Count);
            foreach (var (actualMerge, expectedMerge) in actual.Merges.Zip(expected.Merges))
                actualMerge.Should().BeEquivalentTo(expectedMerge);

            actual.Pictures.Should().HaveCount(expected.Pictures.Count);
            foreach (var (actualMerge, expectedMerge) in actual.Pictures.Zip(expected.Pictures))
                actualMerge.Should().BeEquivalentTo(expectedMerge);
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
using System.Globalization;
using Gooseberry.ExcelStreaming.Tests.Excel;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class ExcelWriterMergeCellsTests
{
    [Fact]
    public async Task ExcelWriterAddCellWithMerges_WritesCorrectData()
    {
        var outputStream = new MemoryStream();

        await using (var writer = new ExcelWriter(outputStream))
        {
            await writer.StartSheet("test sheet");

            await writer.StartRow();
            writer.AddCell("Id", rightMerge: 1, downMerge: 1);

            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var sheets = ExcelReader.ReadSheets(outputStream);

        var expectedSheet = new Excel.Sheet(
            "test sheet",
            new[]
            {
                new Row(new[]
                {
                    new Cell("Id", CellValueType.String)
                })
            },
            Merges: new[] { "A1:B2" });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }

    [Fact]
    public async Task ExcelWriterAddManyCellsWithMerges_WritesCorrectData()
    {
        var outputStream = new MemoryStream();

        var now = DateTime.Now;

        await using (var writer = new ExcelWriter(outputStream))
        {
            await writer.StartSheet("test sheet");

            await writer.StartRow();
            writer.AddCell("Id", rightMerge: 1, downMerge: 1);
            writer.AddEmptyCell();

            writer.AddCell("Dates", rightMerge: 1);
            writer.AddEmptyCell();

            await writer.StartRow();
            writer.AddEmptyCell();
            writer.AddEmptyCell();
            writer.AddCell("Create");
            writer.AddCell("Update");

            await writer.StartRow();
            writer.AddCell(1, rightMerge: 1);
            writer.AddEmptyCell();
            writer.AddCell(now);
            writer.AddCell(now);

            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var sheets = ExcelReader.ReadSheets(outputStream);

        var expectedSheet = new Excel.Sheet(
            "test sheet",
            new[]
            {
                new Row(new[]
                {
                    new Cell("Id", CellValueType.String),
                    new Cell("", CellValueType.String),
                    new Cell("Dates", CellValueType.String),
                    new Cell("", CellValueType.String),
                }),
                new Row(new[]
                {
                    new Cell("", CellValueType.String),
                    new Cell("", CellValueType.String),
                    new Cell("Create", CellValueType.String),
                    new Cell("Update", CellValueType.String)
                }),
                new Row(new[]
                {
                    new Cell("1", CellValueType.Number, Constants.DefaultNumberStyle),
                    new Cell("", CellValueType.String),
                    new Cell(now.ToOADate().ToString(CultureInfo.InvariantCulture), Style: Constants.DefaultDateTimeStyle),
                    new Cell(now.ToOADate().ToString(CultureInfo.InvariantCulture), Style: Constants.DefaultDateTimeStyle)
                })
            },
            Merges: new[] { "A1:B2", "C1:D1", "A3:B3" });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }
}
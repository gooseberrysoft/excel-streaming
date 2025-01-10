using System.Globalization;
using Gooseberry.ExcelStreaming.Tests.Excel;
using Color = System.Drawing.Color;
using Gooseberry.ExcelStreaming.Styles;
using Xunit;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class ExcelWriterStylesTests
{
    [Fact]
    public async Task ExcelWriter_WritesCorrectDefaultStyles()
    {
        var outputStream = new MemoryStream();

        await using (var writer = new ExcelWriter(outputStream))
        {
            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var styles = ExcelReader.ReadStyles(outputStream);
        styles.ShouldBeEquivalentTo(Constants.DefaultNumberStyle, Constants.DefaultDateTimeStyle, Constants.DefaultHyperlinkStyle);
    }

    [Fact]
    public async Task AddCellWithStyle_WritesCorrectData()
    {
        var style = new Style(
            Format: "General",
            Font: new Font(Size: 11, Name: "test", Color: Color.Black, Bold: true, Italic: true, Strike: true, Underline: Underline.Single),
            Borders: new Borders(
                Left: new Border(Style: BorderStyle.Thin, Color: Color.Gray),
                Right: new Border(Style: BorderStyle.Thin, Color: Color.Gray),
                Top: new Border(Style: BorderStyle.Thin, Color: Color.Gray),
                Bottom: new Border(Style: BorderStyle.Thin, Color: Color.Gray)),
            Fill: new Fill(Pattern: FillPattern.Solid, Color: Color.Aqua),
            Alignment: new Alignment(Horizontal: HorizontalAlignment.Center, Vertical: VerticalAlignment.Top, WrapText: true));

        var outputStream = new MemoryStream();

        var now = DateTime.Now;

        var styles = new StylesSheetBuilder();
        var styleReference = styles.GetOrAdd(style);
        await using (var writer = new ExcelWriter(outputStream, styles.Build()))
        {
            await writer.StartSheet("test sheet");

            await writer.StartRow();

            writer.AddCell("text", styleReference);
            writer.AddCell(1, styleReference);
            writer.AddCell(2L, styleReference);
            writer.AddCell(3.55m, styleReference);
            writer.AddCell(now, styleReference);
            writer.AddEmptyCell(styleReference);

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
                    new Cell("text", CellValueType.String, style),
                    new Cell("1", CellValueType.Number, style),
                    new Cell("2", CellValueType.Number, style),
                    new Cell("3.55", CellValueType.Number, style),
                    new Cell(now.ToInternalOADate(), Style: style),
                    new Cell("", CellValueType.String, style)
                })
            });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }

    [Fact]
    public async Task AddCellWithFormat_WritesCorrectData()
    {
        var style = new Style(
            Format: StandardFormat.DayMonthYear4WithSlashes,
            Borders: new Borders(
                Left: new Border(Style: BorderStyle.Thin, Color: Color.Gray),
                Right: new Border(Style: BorderStyle.Thin, Color: Color.Gray),
                Top: new Border(Style: BorderStyle.Thin, Color: Color.Gray),
                Bottom: new Border(Style: BorderStyle.Thin, Color: Color.Gray)),
            Font: new Font(Size: 11, Name: "test", Color: Color.Black, Bold: true, Italic: true, Strike: true, Underline: Underline.Single),
            Fill: new Fill(Pattern: FillPattern.Solid, Color: Color.Aqua));

        var outputStream = new MemoryStream();

        var now = DateTime.Now;
        var today = DateOnly.FromDateTime(now);

        var styles = new StylesSheetBuilder();
        var styleReference = styles.GetOrAdd(style);
        await using (var writer = new ExcelWriter(outputStream, styles.Build()))
        {
            await writer.StartSheet("test sheet");

            await writer.StartRow();

            writer.AddCell("text", styleReference);
            writer.AddCell(1, styleReference);
            writer.AddCell(2L, styleReference);
            writer.AddCell(3.55m, styleReference);
            writer.AddCell(now, styleReference);
            writer.AddCell(today, styleReference);
            writer.AddEmptyCell(styleReference);

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
                    new Cell("text", CellValueType.String, style),
                    new Cell("1", CellValueType.Number, style),
                    new Cell("2", CellValueType.Number, style),
                    new Cell("3.55", CellValueType.Number, style),
                    new Cell(now.ToInternalOADate(), Style: style),
                    new Cell(today.ToDateTime(default).ToInternalOADate(), Style: style),
                    new Cell("", CellValueType.String, style)
                })
            });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }
}
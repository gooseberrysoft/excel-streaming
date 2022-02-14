using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using Gooseberry.ExcelStreaming.Tests.Excel;
using Color = System.Drawing.Color;
using Gooseberry.ExcelStreaming.Styles;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests
{
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
            styles.ShouldBeEquivalentTo(Constants.DefaultNumberStyle, Constants.DefaultDateTimeStyle);
        }

        [Fact]
        public async Task AddCellWithStyle_WritesCorrectData()
        {
            var style = new Style(
                format: "General",
                font: new Font(size: 11, name: "test", color: Color.Black, bold: true, italic: true, strike: true, underline: Underline.Single),
                borders: new Borders(
                    left: new Border(style: BorderStyle.Thin, color: Color.Gray),
                    right: new Border(style: BorderStyle.Thin, color: Color.Gray),
                    top: new Border(style: BorderStyle.Thin, color: Color.Gray),
                    bottom: new Border(style: BorderStyle.Thin, color: Color.Gray)),
                fill: new Fill(pattern: FillPattern.Solid, color: Color.Aqua),
                alignment: new Alignment(horizontal: HorizontalAlignment.Center, vertical: VerticalAlignment.Top, wrapText: true));

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
            using (var file = new FileStream("test.xlsx", FileMode.Create))
                await outputStream.CopyToAsync(file);
            outputStream.Seek(0, SeekOrigin.Begin);

            var sheets = ExcelReader.ReadSheets(outputStream);

            var expectedSheet = new Sheet(
                "test sheet",
                new []
                {
                    new Row(new []
                    {
                        new Cell("text", CellValueType.String, style),
                        new Cell("1", CellValueType.Number, style),
                        new Cell("2", CellValueType.Number, style),
                        new Cell("3.55", CellValueType.Number, style),
                        new Cell(now.ToOADate().ToString(CultureInfo.InvariantCulture), style: style),
                        new Cell("", CellValueType.String, style)
                    })
                });

            sheets.ShouldBeEquivalentTo(expectedSheet);
        }
    }
}

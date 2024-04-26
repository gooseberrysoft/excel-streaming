using System.Globalization;
using System.Text;
using Gooseberry.ExcelStreaming.Tests.Excel;
using FluentAssertions;
using Gooseberry.ExcelStreaming.Configuration;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class ExcelWriterTests
{
    [Fact]
    public async Task ExcelWriter_WritesCorrectData()
    {
        var outputStream = new MemoryStream();

        var now = DateTime.Now;

        await using (var writer = new ExcelWriter(outputStream))
        {
            await writer.StartSheet("test sheet");

            await writer.StartRow();
            writer.AddCell("Id");
            writer.AddCell("Name");
            writer.AddCell("Date");

            await writer.StartRow();
            writer.AddCell(1);
            writer.AddCell("name");
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
                    new Cell("Name", CellValueType.String),
                    new Cell("Date", CellValueType.String)
                }),
                new Row(new[]
                {
                    new Cell("1", CellValueType.Number, Constants.DefaultNumberStyle),
                    new Cell("name", CellValueType.String),
                    new Cell(now.ToOADate().ToString(CultureInfo.InvariantCulture), style: Constants.DefaultDateTimeStyle)
                })
            });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }

    [Fact]
    public async Task ExcelWriterTwoSheets_WritesCorrectData()
    {
        using var outputStream = new MemoryStream();

        await using (var writer = new ExcelWriter(outputStream))
        {
            await writer.StartSheet("test sheet 1");

            await writer.StartRow();
            writer.AddCell("Id");
            writer.AddCell("Name");

            await writer.StartRow();
            writer.AddCell(1);
            writer.AddCell("name");

            await writer.StartSheet("test sheet 2");

            await writer.StartRow();
            writer.AddCell("Number");
            writer.AddCell("Definition");

            await writer.StartRow();
            writer.AddCell(3);
            writer.AddCell("three");

            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var sheets = ExcelReader.ReadSheets(outputStream);

        var expectedSheets = new[]
        {
            new Excel.Sheet(
                "test sheet 1",
                new[]
                {
                    new Row(new[]
                    {
                        new Cell("Id", CellValueType.String),
                        new Cell("Name", CellValueType.String)
                    }),
                    new Row(new[]
                    {
                        new Cell("1", CellValueType.Number, Constants.DefaultNumberStyle),
                        new Cell("name", CellValueType.String)
                    })
                }
            ),
            new Excel.Sheet(
                "test sheet 2",
                new[]
                {
                    new Row(new[]
                    {
                        new Cell("Number", CellValueType.String),
                        new Cell("Definition", CellValueType.String)
                    }),
                    new Row(new[]
                    {
                        new Cell("3", CellValueType.Number, Constants.DefaultNumberStyle),
                        new Cell("three", CellValueType.String)
                    })
                }
            ),
        };

        sheets.ShouldBeEquivalentTo(expectedSheets);
    }

    [Fact]
    public async Task ExcelWriterEmptyDocument_WritesCorrectData()
    {
        var outputStream = new MemoryStream();

        await using (var writer = new ExcelWriter(outputStream))
        {
            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var sheets = ExcelReader.ReadSheets(outputStream);

        sheets.Should().BeEmpty();
    }

    [Theory]
    [MemberData(nameof(SpecialSymbols))]
    public async Task AddCellWithSpecialSymbols_WritesCorrectData(string value)
    {
        var outputStream = new MemoryStream();

        var longText = string.Join('&', Enumerable.Repeat(0, 2000)
            .Select(_ => value));
        longText = longText.Substring(0, Math.Min(longText.Length, 32_767));

        await using (var writer = new ExcelWriter(outputStream, bufferSize: 1024))
        {
            await writer.StartSheet("test");

            await writer.StartRow();
            writer.AddCell(value);
            writer.AddCell(longText);

            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var sheets = ExcelReader.ReadSheets(outputStream);

        var expectedSheet = new Excel.Sheet(
            "test",
            new[]
            {
                new Row(new[]
                {
                    new Cell(value, CellValueType.String),
                    new Cell(longText, CellValueType.String),
                })
            });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }


    [Fact]
    public async Task AddCellsWithSpecialSymbols_WritesCorrectData()
    {
        const string text = "Tags such as <img> and <input> directly &introduce \"content\" into the page.";
        const string text2 = "<img> is image tag";
        const string text3 = "Exit >>>>>>>>>>>>>>>";

        var longTextBuilder = new StringBuilder(32_767);
        for (var i = 0; i < 100; i++)
            longTextBuilder.Append(text).Append(text2).Append(text3);

        var longText = longTextBuilder.ToString();

        var outputStream = new MemoryStream();

        await using (var writer = new ExcelWriter(outputStream))
        {
            await writer.StartSheet("test");

            await writer.StartRow();
            writer.AddCell(text);
            writer.AddCell(text2);
            writer.AddCell(text3);
            writer.AddCell(longText);

            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var sheets = ExcelReader.ReadSheets(outputStream);

        var expectedSheet = new Excel.Sheet(
            "test",
            new[]
            {
                new Row(new[]
                {
                    new Cell(text, CellValueType.String),
                    new Cell(text2, CellValueType.String),
                    new Cell(text3, CellValueType.String),
                    new Cell(longText, CellValueType.String),
                })
            });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }

    [Fact]
    public async Task AddCellsWithSpecialSymbolsUtf8_WritesCorrectData()
    {
        var text = "Tags such as <img> and <input> directly &introduce \"content\" into the page.";
        var text2 = "<img> is image tag";
        var text3 = "Exit >>>>>>>>>>>>>>>";

        var outputStream = new MemoryStream();

        await using (var writer = new ExcelWriter(outputStream))
        {
            await writer.StartSheet("test");

            await writer.StartRow();

            WriteUtf8Row(writer);

            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var sheets = ExcelReader.ReadSheets(outputStream);

        var expectedSheet = new Excel.Sheet(
            "test",
            new[]
            {
                new Row(new[]
                {
                    new Cell(text, CellValueType.String),
                    new Cell(text2, CellValueType.String),
                    new Cell(text3, CellValueType.String),
                })
            });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }

    private static void WriteUtf8Row(ExcelWriter writer)
    {
        var text = "Tags such as <img> and <input> directly &introduce \"content\" into the page."u8;
        var text2 = "<img> is image tag"u8;
        var text3 = "Exit >>>>>>>>>>>>>>>"u8;

        writer.AddUtf8Cell(text);
        writer.AddUtf8Cell(text2);
        writer.AddUtf8Cell(text3);
    }

    [Theory]
    [MemberData(nameof(SpecialSymbols))]
    public async Task StartSheetWithSpecialSymbols_WritesCorrectData(string sheetName)
    {
        var outputStream = new MemoryStream();

        await using (var writer = new ExcelWriter(outputStream))
        {
            await writer.StartSheet(sheetName);

            await writer.StartRow();
            writer.AddCell("test");

            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var sheets = ExcelReader.ReadSheets(outputStream);

        var expectedSheet = new Excel.Sheet(
            sheetName,
            new[]
            {
                new Row(new[]
                {
                    new Cell("test", CellValueType.String),
                })
            });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }

    [Fact]
    public async Task AddCellWithDataLongerThanBuffer_WritesCorrectData()
    {
        var outputStream = new MemoryStream();
        var longString = "long long long loong loooong loooooooon loooooooooooooooooooooong very long string";

        await using (var writer = new ExcelWriter(outputStream, bufferSize: 32))
        {
            await writer.StartSheet("test");

            await writer.StartRow();
            writer.AddCell(longString);

            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var sheets = ExcelReader.ReadSheets(outputStream);

        var expectedSheet = new Excel.Sheet(
            "test",
            new[]
            {
                new Row(new[]
                {
                    new Cell(longString, CellValueType.String),
                })
            });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }

    [Fact]
    public async Task StartSheetWithNameLongerThanBuffer_WritesCorrectData()
    {
        var outputStream = new MemoryStream();
        var longString =
            "long long long loong loooong loooooooon loooooooooooooooooooooong very long" +
            "long long long loong loooong loooooooon loooooooooooooooooooooong very long &string";

        await using (var writer = new ExcelWriter(outputStream, bufferSize: 32))
        {
            await writer.StartSheet(longString);

            await writer.StartRow();
            writer.AddCell("test");

            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var sheets = ExcelReader.ReadSheets(outputStream);

        var expectedSheet = new Excel.Sheet(
            longString,
            new[]
            {
                new Row(new[]
                {
                    new Cell("test", CellValueType.String),
                })
            });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }

    [Fact]
    public async Task StartSheet_WritesCorrectColumnWidths()
    {
        using var outputStream = new MemoryStream();

        await using (var writer = new ExcelWriter(outputStream))
        {
            await writer.StartSheet(
                "test sheet",
                new SheetConfiguration(
                    new[]
                    {
                        new Column(width: 10m),
                        new Column(width: 15m)
                    }));

            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var sheets = ExcelReader.ReadSheets(outputStream);

        var expectedSheet = new Excel.Sheet(
            "test sheet",
            Array.Empty<Row>(),
            new[] { new Column(10m), new Column(15m) });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }

    [Fact]
    public async Task AddRow_WritesCorrectRowHeight()
    {
        using var outputStream = new MemoryStream();

        await using (var writer = new ExcelWriter(outputStream))
        {
            await writer.StartSheet("test sheet");
            await writer.StartRow(10.8m);

            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);

        var sheets = ExcelReader.ReadSheets(outputStream);

        var expectedSheet = new Excel.Sheet(
            "test sheet",
            new[] { new Row(Array.Empty<Cell>(), height: 10.8m) });

        sheets.ShouldBeEquivalentTo(expectedSheet);
    }

    public static IEnumerable<object[]> SpecialSymbols()
    {
        return new[]
        {
            new[] { "\"" },
            new[] { "'" },
            new[] { "<" },
            new[] { ">" },
            new[] { "&" },
            new[] { "text before_\"" },
            new[] { "\"_text after" },
            new[] { "text before_\"_text after" },
            new[] { "test\nother" },
            new[] { "test\tother" },
            new[] { "2021 \u00a9 ozon" }
        };
    }
}
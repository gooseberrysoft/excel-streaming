using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Gooseberry.ExcelStreaming.Tests.Excel;
using FluentAssertions;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests
{
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

            var expectedSheet = new Sheet(
                "test sheet",
                new []
                {
                    new Row(new []
                    {
                        new Cell("Id", CellValueType.String),
                        new Cell("Name", CellValueType.String),
                        new Cell("Date", CellValueType.String)
                    }),
                    new Row(new []
                    {
                        new Cell("1", CellValueType.Number, Constants.DefaultNumberStyle),
                        new Cell("name", CellValueType.String),
                        new Cell(now.ToOADate().ToString(), style: Constants.DefaultDateTimeStyle)
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
                new Sheet(
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
                new Sheet(
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

            await using (var writer = new ExcelWriter(outputStream))
            {
                await writer.StartSheet("test");

                await writer.StartRow();
                writer.AddCell(value);

                await writer.Complete();
            }

            outputStream.Seek(0, SeekOrigin.Begin);

            var sheets = ExcelReader.ReadSheets(outputStream);

            var expectedSheet = new Sheet(
                "test",
                new []
                {
                    new Row(new []
                    {
                        new Cell(value, CellValueType.String),
                    })
                });

            sheets.ShouldBeEquivalentTo(expectedSheet);
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

            var expectedSheet = new Sheet(
                sheetName,
                new []
                {
                    new Row(new []
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

            await using (var writer = new ExcelWriter(outputStream, bufferSize: 8))
            {
                await writer.StartSheet("test");

                await writer.StartRow();
                writer.AddCell("long long very long string");

                await writer.Complete();
            }

            outputStream.Seek(0, SeekOrigin.Begin);

            var sheets = ExcelReader.ReadSheets(outputStream);

            var expectedSheet = new Sheet(
                "test",
                new []
                {
                    new Row(new []
                    {
                        new Cell("long long very long string", CellValueType.String),
                    })
                });

            sheets.ShouldBeEquivalentTo(expectedSheet);
        }

        [Fact]
        public async Task StartSheetWithNameLongerThanBuffer_WritesCorrectData()
        {
            var outputStream = new MemoryStream();

            await using (var writer = new ExcelWriter(outputStream, bufferSize: 8))
            {
                await writer.StartSheet("long long very long string");

                await writer.StartRow();
                writer.AddCell("test");

                await writer.Complete();
            }

            outputStream.Seek(0, SeekOrigin.Begin);

            var sheets = ExcelReader.ReadSheets(outputStream);

            var expectedSheet = new Sheet(
                "long long very long string",
                new []
                {
                    new Row(new []
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
                    new Column(width: 10m),
                    new Column(width: 15m));

                await writer.Complete();
            }

            outputStream.Seek(0, SeekOrigin.Begin);

            var sheets = ExcelReader.ReadSheets(outputStream);

            var expectedSheet = new Sheet(
                "test sheet",
                Array.Empty<Row>(),
                new []{new Column(10m), new Column(15m)});

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

            var expectedSheet = new Sheet(
                "test sheet",
                new [] {new Row(Array.Empty<Cell>(), height: 10.8m)});

            sheets.ShouldBeEquivalentTo(expectedSheet);
        }

        public static IEnumerable<object[]> SpecialSymbols()
        {
            return new[]
            {
                new [] {"\""},
                new [] {"'"},
                new [] {"<"},
                new [] {">"},
                new [] {"&"},
                new [] {"text before_\""},
                new [] {"\"_text after"},
                new [] {"text before_\"_text after"},
                new [] {"test\nother"},
                new [] {"test\tother"},
                new [] {"2021 \u00a9 ozon"}
            };
        }
    }
}

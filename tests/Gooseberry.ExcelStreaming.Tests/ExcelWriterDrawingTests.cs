using System.Drawing;
using Gooseberry.ExcelStreaming.Pictures;
using Gooseberry.ExcelStreaming.Tests.Cases;
using Gooseberry.ExcelStreaming.Tests.Excel;
using Gooseberry.ExcelStreaming.Tests.Extensions;
using Xunit;
using Picture = Gooseberry.ExcelStreaming.Tests.Excel.Picture;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class ExcelWriterDrawingTests
{
    [Theory]
    [MemberData(nameof(ImageCases.GetCases), MemberType = typeof(ImageCases))]
    public async Task OneCellAnchor(ImageCase imageCase)
    {
        var placement1 = new PicturePlacement(
            new AnchorCell(row: 2, column: 3),
            new Size(width: 100, height: 200));

        var placement2 = new PicturePlacement(
            new AnchorCell(row: 6, column: 6, columnOffset: 10, rowOffset: 30),
            new Size(width: 150, height: 250));

        await using var imageStream = imageCase.OpenStream();

        var expectedSheet1 = new Excel.Sheet(
            "test sheet 1",
            new[]
            {
                new Row(
                    new[]
                    {
                        new Cell("Id", CellValueType.String),
                        new Cell("Name", CellValueType.String),
                        new Cell("Date", CellValueType.String)
                    })
            },
            Pictures: new[]
            {
                new Picture(
                    imageStream.ToArray(),
                    placement1 with
                    {
                        Size = new Size(
                            (int)EmuConverter.ConvertToEnglishMetricUnits(placement1.Size!.Value.Width, resolution: 96),
                            (int)EmuConverter.ConvertToEnglishMetricUnits(placement1.Size!.Value.Height, resolution: 96))
                    },
                    imageCase.Format),
                new Picture(
                    imageStream.ToArray(),
                    placement2 with
                    {
                        Size = new Size(
                            (int)EmuConverter.ConvertToEnglishMetricUnits(placement2.Size!.Value.Width, resolution: 96),
                            (int)EmuConverter.ConvertToEnglishMetricUnits(placement2.Size!.Value.Height, resolution: 96))
                    },
                    imageCase.Format)
            });

        var expectedSheet2 = expectedSheet1 with { Name = "test sheet 2" };
        var expectedSheets = new[] { expectedSheet1, expectedSheet2 };

        var outputStream = new MemoryStream();

        await using (var writer = new ExcelWriter(outputStream))
        {
            for (var sheet = 1; sheet <= 2; sheet++)
            {
                await writer.StartSheet($"test sheet {sheet}");

                await writer.StartRow();
                writer.AddCell("Id");
                writer.AddCell("Name");
                writer.AddCell("Date");

                writer.AddPicture(imageStream, imageCase.Format, placement1.From, placement1.Size!.Value);
                writer.AddPicture(imageStream, imageCase.Format, placement2.From, placement2.Size!.Value);
            }
        }

        outputStream.Seek(offset: 0, SeekOrigin.Begin);

        var actualSheets = ExcelReader.ReadSheets(outputStream);

        actualSheets.ShouldBeEquivalentTo(expectedSheets);
    }

    [Theory]
    [MemberData(nameof(ImageCases.GetCases), MemberType = typeof(ImageCases))]
    public async Task TwoCellAnchor(ImageCase imageCase)
    {
        var placement1 = new PicturePlacement(
            new AnchorCell(row: 2, column: 3),
            new AnchorCell(row: 7, column: 10));

        var placement2 = new PicturePlacement(
            new AnchorCell(column: 6, row: 6, columnOffset: 10, rowOffset: 30),
            new AnchorCell(column: 10, row: 7, columnOffset: 15, rowOffset: 300));

        await using var imageStream = imageCase.OpenStream();

        var expectedSheet1 = new Excel.Sheet(
            "test sheet 1",
            new[]
            {
                new Row(
                    new[]
                    {
                        new Cell("Id", CellValueType.String),
                        new Cell("Name", CellValueType.String),
                        new Cell("Date", CellValueType.String)
                    })
            },
            Pictures: new[]
            {
                new Picture(
                    imageStream.ToArray(),
                    placement1,
                    imageCase.Format),
                new Picture(
                    imageStream.ToArray(),
                    placement2,
                    imageCase.Format)
            });

        var expectedSheet2 = expectedSheet1 with { Name = "test sheet 2" };
        var expectedSheets = new[] { expectedSheet1, expectedSheet2 };

        var outputStream = new MemoryStream();

        await using (var writer = new ExcelWriter(outputStream))
        {
            for (var sheet = 1; sheet <= 2; sheet++)
            {
                await writer.StartSheet($"test sheet {sheet}");

                await writer.StartRow();
                writer.AddCell("Id");
                writer.AddCell("Name");
                writer.AddCell("Date");

                writer.AddPicture(imageStream, imageCase.Format, placement1.From, placement1.To!.Value);
                writer.AddPicture(imageStream, imageCase.Format, placement2.From, placement2.To!.Value);
            }
        }

        outputStream.Seek(offset: 0, SeekOrigin.Begin);

        var actualSheets = ExcelReader.ReadSheets(outputStream);

        actualSheets.ShouldBeEquivalentTo(expectedSheets);
    }
}
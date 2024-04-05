using System.Drawing;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using Gooseberry.ExcelStreaming.Pictures.Abstractions;
using Gooseberry.ExcelStreaming.Pictures.InfoReaders;
using Gooseberry.ExcelStreaming.Pictures.Placements;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class ExcelWriterDrawingTests
{
    [Fact]
    public async Task Load()
    {
        await using var stream = Assembly
            .GetExecutingAssembly()
            .GetManifestResourceStream("Gooseberry.ExcelStreaming.Tests.Resources.Images.SampleImagePng.png")!;

        await using var outputStream = new FileStream("test.xlsx", FileMode.Create);

        await using var writer = new ExcelWriter(outputStream);

        await writer.StartSheet("Test sheet");

        await writer.StartRow();
        writer.AddCell("Test col 1");
        writer.AddCell("Test col 2");

        await writer.StartRow();

        writer.AddPicture(
            stream,
            new OneCellAnchorPicturePlacement(
                new AnchorCell(row: 3, column: 2),
                new Size(width: 300, height: 400)));

        await writer.Complete();
    }

    [Fact]
    public async Task Load2()
    {
        await using var stream = Assembly
            .GetExecutingAssembly()
            .GetManifestResourceStream("Gooseberry.ExcelStreaming.Tests.Resources.Images.SampleImagePng.png")!;

        await using (var outputStream = new FileStream("test.xlsx", FileMode.Create))
        {
            await using var writer = new ExcelWriter(outputStream);

            await writer.StartSheet("Test sheet");

            await writer.StartRow();
            writer.AddCell("Test col 1");
            writer.AddCell("Test col 2");

            await writer.StartRow();

            writer.AddPicture(
                stream,
                new TwoCellAnchorPicturePlacement(
                    From: new AnchorCell(row: 3, column: 2),
                    To: new AnchorCell(row: 10, column: 8)));

            await writer.Complete();
        }

        var validator = new OpenXmlValidator();

        var errors = validator.Validate(SpreadsheetDocument.Open("test.xlsx", isEditable: false)).ToArray();
    }
}

using System.Drawing;
using System.Reflection;
using Gooseberry.ExcelStreaming.Tests.ExternalZip;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class ExcelFilesGenerator
{
    private const string skip = "Null me for manual run";
    const string BasePath = "c:\\temp\\excelWriter\\";

    private static readonly Stream Picture = Assembly.GetExecutingAssembly().GetManifestResourceStream(
        "Gooseberry.ExcelStreaming.Tests.Resources.Images.Humpty_Dumpty.jpg")!;

    [Fact(Skip = skip)]
    public async Task Basic()
    {
        await using var outputStream = new FileStream(BasePath + "Basic.xlsx", FileMode.Create);

        await GenerateContent(outputStream);
    }

    [Fact(Skip = skip)]
    public async Task Basic_SharpZipLib()
    {
        await using var outputStream = new FileStream(BasePath + "Basic_SharpZipLib.xlsx", FileMode.Create);

        await GenerateContent(archive: new SharpZipLibArchive(outputStream));
    }

    [Fact(Skip = skip)]
    public async Task Basic_SharpCompress()
    {
        await using var outputStream = new FileStream(BasePath + "Basic_SharpCompress.xlsx", FileMode.Create);

        await GenerateContent(archive: new SharpCompressZipArchive(outputStream));
    }

    private static async Task GenerateContent(Stream? stream = null, IZipArchive? archive = null)
    {
        var sharedStringTableBuilder = new SharedStringTableBuilder();
        var sharedStringRef1 = sharedStringTableBuilder.GetOrAdd(">>> Shared string with special & symbols <<< ");
        var sharedStringRef2 = sharedStringTableBuilder.GetOrAdd("“Tell us a story!” said the March Hare.");
        var sharedStringTable = sharedStringTableBuilder.Build();

        await using var writer = stream != null
            ? new ExcelWriter(stream!, sharedStringTable: sharedStringTable)
            : new ExcelWriter(archive!, sharedStringTable: sharedStringTable);

        for (var sheetIndex = 0; sheetIndex < 3; sheetIndex++)
        {
            await writer.StartSheet($"test#{sheetIndex}");

            for (var row = 0; row < 10_000; row++)
            {
                await writer.StartRow();

                for (var columnBatch = 0; columnBatch < 2; columnBatch++)
                {
                    writer.AddCell(row);
                    writer.AddCell(DateTime.Now.Ticks);
                    writer.AddCell(DateTime.Now);
                    writer.AddCell(1234567.9876M);
                    writer.AddCell("Tags such as <img> and <input> directly introduce content into the page.");
                    writer.AddCell("The cat (Felis catus), commonly referred to as the domestic cat");
                    writer.AddCellWithSharedString(
                        "The dog (Canis familiaris or Canis lupus familiaris) is a domesticated descendant of the wolf");
                    writer.AddUtf8Cell("Utf8 string with <tags>"u8);
                    writer.AddCell("String as chars".AsSpan());
                    writer.AddCell(new Hyperlink("https://github.com/gooseberrysoft/excel-streaming", "Excel Streaming"));
                    writer.AddCell(sharedStringRef1);
                    writer.AddCell(sharedStringRef2);
                }
            }

            writer.AddPicture(Picture, PictureFormat.Jpeg, new AnchorCell(3, 10_001), new Size(100, 130));
            writer.AddPicture(Picture, PictureFormat.Jpeg, new AnchorCell(10, 10_001), new AnchorCell(15, 10_010));
        }

        await writer.Complete();
    }

    [Fact(Skip = skip)]
    public async Task Primitive()
    {
        await using var outputStream = new FileStream(BasePath + "Primitive.xlsx", FileMode.Create);

        await using var writer = new ExcelWriter(outputStream);

        for (var sheetIndex = 0; sheetIndex < 3; sheetIndex++)
        {
            await writer.StartSheet($"test#{sheetIndex}");

            for (var row = 0; row < 1_000; row++)
            {
                await writer.StartRow();

                for (var columnBatch = 0; columnBatch < 2; columnBatch++)
                {
                    writer.AddCell(row);
                    writer.AddCell(DateTime.Now.Ticks);
                    writer.AddCell(DateTime.Now);
                    writer.AddCell(1234567.9876M);
                    writer.AddCell("Tags such as <img> and <input> directly introduce content into the page.");
                    writer.AddUtf8Cell("Utf8 string with <tags>"u8);
                    writer.AddCell("String as chars".AsSpan());
                }
            }
        }


        await writer.Complete();
    }
}
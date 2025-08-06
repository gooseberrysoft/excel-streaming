using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class ContentTypesWriter
{
    private static ReadOnlySpan<byte> Prefix => "<Override PartName=\"/xl/worksheets/sheet"u8;

    private static ReadOnlySpan<byte> Postfix =>
        ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"u8;

    private static ReadOnlySpan<byte> TypesPrefix =>
        """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
         <Default Extension ="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
         <Override PartName ="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
         <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml" />
        """u8;

    private static ReadOnlySpan<byte> TypesPostfix => "</Types>"u8;

    private static ReadOnlySpan<byte> DrawingPrefix => "<Override PartName=\"/"u8;

    private static ReadOnlySpan<byte> DrawingPostfix => "\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>"u8;

    public static void Write(IReadOnlyCollection<Sheet> sheets, SheetDrawings sheetDrawings, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);
        TypesPrefix.WriteTo(buffer, ref span, ref written);

        foreach (var pictureFormat in sheetDrawings.Pictures.Select(v => v.Format).Distinct())
            Get(pictureFormat).WriteTo(buffer, ref span, ref written);

        foreach (var sheet in sheets)
        {
            Prefix.WriteTo(buffer, ref span, ref written);
            sheet.Id.WriteTo(buffer, ref span, ref written);
            Postfix.WriteTo(buffer, ref span, ref written);

            var drawing = sheetDrawings.Get(sheet.Id);

            if (!drawing.IsEmpty)
            {
                DrawingPrefix.WriteTo(buffer, ref span, ref written);
                PathResolver.GetDrawingFullPath(drawing).WriteTo(buffer, encoder, ref span, ref written);
                DrawingPostfix.WriteTo(buffer, ref span, ref written);
            }
        }

        TypesPostfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }

    private static ReadOnlySpan<byte> Get(PictureFormat format)
    {
        return format switch
        {
            PictureFormat.Bmp => GetBmp(),
            PictureFormat.Gif => GetGif(),
            PictureFormat.Png => GetPng(),
            PictureFormat.Tiff => GetTiff(),
            PictureFormat.Icon => GetIcon(),
            PictureFormat.Jpeg => GetJpeg(),
            PictureFormat.Emf => GetEmf(),
            PictureFormat.Wmf => GetWmf(),
            _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unknown picture format")
        };
    }

    private static ReadOnlySpan<byte> GetBmp()
        => "<Default Extension=\"bmp\" ContentType=\"image/bmp\"/>"u8;

    private static ReadOnlySpan<byte> GetGif()
        => "<Default Extension=\"gif\" ContentType=\"image/gif\"/>"u8;

    private static ReadOnlySpan<byte> GetPng()
        => "<Default Extension=\"png\" ContentType=\"image/png\"/>"u8;

    private static ReadOnlySpan<byte> GetTiff()
        => "<Default Extension=\"tiff\" ContentType=\"image/tiff\"/>"u8;

    private static ReadOnlySpan<byte> GetIcon()
        => "<Default Extension=\"icon\" ContentType=\"image/x-icon\"/>"u8;

    private static ReadOnlySpan<byte> GetJpeg()
        => "<Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>"u8;

    private static ReadOnlySpan<byte> GetEmf()
        => "<Default Extension=\"emf\" ContentType=\"image/x-emf\"/>"u8;

    private static ReadOnlySpan<byte> GetWmf()
        => "<Default Extension=\"wmf\" ContentType=\"image/x-wmf\"/>"u8;
}
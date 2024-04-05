using Gooseberry.ExcelStreaming.Pictures;
using Gooseberry.ExcelStreaming.Pictures.Abstractions;

namespace Gooseberry.ExcelStreaming;

internal static class PathResolver
{
    private const string MainDirectory = "xl";

    public static string GetFullPath(in Sheet sheet)
        => $"{MainDirectory}/{GetDirectory(sheet)}/{GetFileName(sheet)}";

    public static string GetRelationshipsFullPath(in Sheet sheet)
        => $"{MainDirectory}/{GetRelationshipsDirectory(sheet)}/{GetRelationshipsFileName(sheet)}";

    public static string GetFullPath(Drawing drawing)
        => $"{MainDirectory}/{GetDirectory(drawing)}/{GetFileName(drawing)}";

    public static string GetRelationshipsFullPath(Drawing drawing)
        => $"{MainDirectory}/{GetRelationshipsDirectory(drawing)}/{GetRelationshipsFileName(drawing)}";

    public static string GetFullPath(Picture picture)
        => $"{MainDirectory}/{GetDirectory(picture)}/{GetFileName(picture)}";

    public static string GetFileName(Picture picture)
        => $"image{picture.Id}.{GetExtension(picture.Format)}";

    public static string GetDirectory(Picture picture)
        => $"media";

    public static string GetFileName(Drawing drawing)
        => $"drawing{drawing.SheetId}.xml";

    public static string GetRelationshipsFileName(Drawing drawing)
        => GetFileName(drawing) + ".rels";

    public static string GetFileName(in Sheet sheet)
        => $"sheet{sheet.Id}.xml";

    public static string GetRelationshipsFileName(in Sheet sheet)
        => GetFileName(in sheet) + ".rels";

    public static string GetDirectory(Drawing drawing)
        => "drawings";

    public static string GetRelationshipsDirectory(Drawing drawing)
        => GetDirectory(drawing) + "/_rels";

    public static string GetDirectory(in Sheet sheet)
        => "worksheets";

    public static string GetRelationshipsDirectory(in Sheet sheet)
        => GetDirectory(sheet) + "/_rels";

    private static string GetExtension(PictureFormat format)
    {
        return format switch
        {
            PictureFormat.Bmp => "bmp",
            PictureFormat.Gif => "gif",
            PictureFormat.Png => "png",
            PictureFormat.Tiff => "tiff",
            PictureFormat.Icon => "ico",
            PictureFormat.Jpeg => "jpg",
            PictureFormat.Emf => "emf",
            PictureFormat.Wmf => "wmf",
            _ => "image",
        };
    }
}

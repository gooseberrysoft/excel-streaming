using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming;

internal static class PathResolver
{
    private const string MainDirectory = "xl";

    public static string GetSheetFullPath(int sheetId)
        => $"{MainDirectory}/{GetSheetDirectory()}/{GetSheetFileName(sheetId)}";

    public static string GetSheetRelationshipsFullPath(int sheetId)
        => $"{MainDirectory}/{GetSheetRelationshipsDirectory()}/{GetSheetRelationshipsFileName(sheetId)}";

    public static string GetDrawingFullPath(Drawing drawing)
        => $"{MainDirectory}/{GetDrawingDirectory()}/{GetDrawingFileName(drawing)}";

    public static string GetDrawingRelationshipsFullPath(Drawing drawing)
        => $"{MainDirectory}/{GetDrawingRelationshipsDirectory()}/{GetDrawingRelationshipsFileName(drawing)}";

    public static string GetPictureFullPath(Picture picture)
        => $"{MainDirectory}/{GetPictureDirectory()}/{GetPictureFileName(picture)}";

    public static string GetPictureFileName(Picture picture)
        => $"image{picture.Id}.{GetExtension(picture.Format)}";

    public static string GetPictureDirectory()
        => "media";

    public static string GetDrawingFileName(Drawing drawing)
        => $"drawing{drawing.SheetId}.xml";

    public static string GetDrawingRelationshipsFileName(Drawing drawing)
        => GetDrawingFileName(drawing) + ".rels";

    public static string GetSheetFileName(int sheetId)
        => $"sheet{sheetId}.xml";

    public static string GetSheetRelationshipsFileName(int sheetId)
        => GetSheetFileName(sheetId) + ".rels";

    public static string GetDrawingDirectory()
        => "drawings";

    public static string GetDrawingRelationshipsDirectory()
        => GetDrawingDirectory() + "/_rels";

    public static string GetSheetDirectory()
        => "worksheets";

    public static string GetSheetRelationshipsDirectory()
        => GetSheetDirectory() + "/_rels";

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

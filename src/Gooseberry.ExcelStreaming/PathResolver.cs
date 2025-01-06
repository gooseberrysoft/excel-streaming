using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming;

internal static class PathResolver
{
    public static string GetSheetFullPath(string sheetName) 
        => $"xl/worksheets/{sheetName}.xml";

    public static string GetSheetRelationshipsFullPath(int sheetId)
        => $"xl/worksheets/_rels/sheet{sheetId}.xml.rels";

    public static string GetDrawingFullPath(Drawing drawing)
        => $"xl/drawings/drawing{drawing.SheetId}.xml";

    public static string GetDrawingRelationshipsFullPath(Drawing drawing)
        => $"xl/drawings/_rels/drawing{drawing.SheetId}.xml.rels";

    public static string GetPictureFullPath(Picture picture)
        => $"xl/media/image{picture.Id}.{GetExtension(picture.Format)}";

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
using Gooseberry.ExcelStreaming.Pictures;

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

    public static string GetRelativePath(in Sheet from, Drawing to)
        => $"../{GetDirectory(to)}/{GetFileName(to)}";

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
        => GetDirectory(drawing) + "/rels";

    public static string GetDirectory(in Sheet sheet)
        => "worksheets";

    public static string GetRelationshipsDirectory(in Sheet sheet)
        => GetDirectory(sheet) + "/_rels";
}

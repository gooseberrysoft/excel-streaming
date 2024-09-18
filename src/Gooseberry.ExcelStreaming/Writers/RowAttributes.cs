namespace Gooseberry.ExcelStreaming.Writers;

public readonly record struct RowAttributes(
    decimal? Height = null,
    byte? OutlineLevel = null,
    bool? IsHidden = false,
    bool? IsCollapsed = false)
{
    public bool IsEmpty()
    {
        return !Height.HasValue &&
               !OutlineLevel.HasValue &&
               (!IsHidden.HasValue || !IsHidden.Value) &&
               (!IsCollapsed.HasValue || !IsCollapsed.Value);
    }
    
    public static RowAttributes Empty => new();
}
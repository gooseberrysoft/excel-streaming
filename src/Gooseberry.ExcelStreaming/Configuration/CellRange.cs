namespace Gooseberry.ExcelStreaming;

public readonly struct CellRange
{
    internal readonly CellReference FromCell;
    internal readonly CellReference ToCell;
    internal readonly string? Range;

    /// <summary>
    /// Cell range, e.g. "A1:C1"
    /// </summary>
    /// <param name="range"></param>
    public CellRange(string range)
    {
        Range = range;
    }

    public CellRange(CellReference fromCell, CellReference toCell)
    {
        FromCell = fromCell;
        ToCell = toCell;
        Range = null;
    }

    public static implicit operator CellRange(string range)
        => new(range);
}
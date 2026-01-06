using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming;

[StructLayout(LayoutKind.Auto)]
internal readonly struct Merge(uint fromRow, uint fromColumn, uint rowSpan, uint colSpan)
{
    public CellReference TopLeft { get; } = new(fromColumn, fromRow);

    public CellReference RightBottom { get; } = new(fromColumn + colSpan, fromRow + rowSpan);
}
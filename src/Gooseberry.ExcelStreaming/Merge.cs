using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming;

[StructLayout(LayoutKind.Auto)]
internal readonly struct Merge(uint fromRow, uint fromColumn, uint downSize, uint rightSize)
{
    public CellReference TopLeft { get; } = new(fromRow, fromColumn);

    public CellReference RightBottom { get; } = new(fromRow + downSize, fromColumn + rightSize);
}
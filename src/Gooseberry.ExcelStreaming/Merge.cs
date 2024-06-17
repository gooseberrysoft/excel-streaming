using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming;

[StructLayout(LayoutKind.Auto)]
internal readonly struct Merge(uint fromRow, uint fromColumn, uint downSize, uint rightSize)
{
    public CellReference TopLeft { get; } = new(fromColumn, fromRow);

    public CellReference RightBottom { get; } = new(fromColumn + rightSize, fromRow + downSize);
}
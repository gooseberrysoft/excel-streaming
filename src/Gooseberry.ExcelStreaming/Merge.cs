using System.Runtime.InteropServices;
using Gooseberry.ExcelStreaming.Configuration;

namespace Gooseberry.ExcelStreaming;

[StructLayout(LayoutKind.Auto)]
internal readonly struct Merge
{
    public Merge(uint fromRow, uint fromColumn, uint downSize, uint rightSize)
    {
            TopLeft = new CellReference(fromRow, fromColumn);
            RightBottom = new CellReference(fromRow + downSize, fromColumn + rightSize);
        }

    public CellReference TopLeft { get; }
        
    public CellReference RightBottom { get; }
}
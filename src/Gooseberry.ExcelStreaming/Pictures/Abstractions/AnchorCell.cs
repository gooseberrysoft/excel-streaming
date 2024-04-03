using System.Drawing;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Pictures.Abstractions;

[StructLayout(LayoutKind.Auto)]
public readonly record struct AnchorCell(int Row, int Column, Point Offset)
{
    public AnchorCell(int row, int column)
        : this(row, column, Point.Empty)
    {
    }
}

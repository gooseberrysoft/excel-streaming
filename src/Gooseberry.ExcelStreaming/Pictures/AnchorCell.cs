using System.Drawing;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Pictures;

[StructLayout(LayoutKind.Auto)]
public readonly record struct AnchorCell(int Column, int Row, Point Offset)
{
    public AnchorCell(int column, int row)
        : this(column, row, Point.Empty)
    {
    }

    public AnchorCell(int column, int row, int columnOffset, int rowOffset)
        : this(column, row, new Point(columnOffset, rowOffset))
    {
    }
}

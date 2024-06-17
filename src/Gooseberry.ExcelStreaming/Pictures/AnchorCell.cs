using System.Drawing;
using System.Runtime.InteropServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

[StructLayout(LayoutKind.Auto)]
public readonly record struct AnchorCell(uint Column, uint Row, Point Offset)
{
    public AnchorCell(uint column, uint row)
        : this(column, row, Point.Empty)
    {
    }

    public AnchorCell(uint column, uint row, int columnOffset, int rowOffset)
        : this(column, row, new Point(columnOffset, rowOffset))
    {
    }
}
using System.Drawing;
using System.Runtime.InteropServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

/// <summary>
/// Creates AnchorCell
/// </summary>
/// <param name="Column">Zero-based column index</param>
/// <param name="Row">Zero-based row index</param>
/// <param name="Offset"></param>
[StructLayout(LayoutKind.Auto)]
public readonly record struct AnchorCell(uint Column, uint Row, Point Offset)
{
    /// <summary>
    /// Creates AnchorCell
    /// </summary>
    /// <param name="column">Zero-based column index</param>
    /// <param name="row">Zero-based row index</param>
    public AnchorCell(uint column, uint row)
        : this(column, row, Point.Empty)
    {
    }

    /// <summary>
    /// Creates AnchorCell
    /// </summary>
    /// <param name="column">Zero-based column index</param>
    /// <param name="row">Zero-based row index</param>
    /// <param name="columnOffset"></param>
    /// <param name="rowOffset"></param>
    public AnchorCell(uint column, uint row, int columnOffset, int rowOffset)
        : this(column, row, new Point(columnOffset, rowOffset))
    {
    }
}
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Tests.Excel;

[StructLayout(LayoutKind.Auto)]
public readonly struct Row : IEquatable<Row>
{
    public Row(
        IReadOnlyCollection<Cell> cells,
        decimal? height = null,
        byte? outlineLevel = null,
        bool? isHidden = null,
        bool? isCollapsed = null)
    {
        Cells = cells;
        Height = height;
        OutlineLevel = outlineLevel;
        IsHidden = isHidden;
        IsCollapsed = isCollapsed;
    }

    public decimal? Height { get; }

    public byte? OutlineLevel { get; }

    public bool? IsHidden { get; }

    public bool? IsCollapsed { get; }

    public IReadOnlyCollection<Cell> Cells { get; }

    public bool Equals(Row other)
        => Cells.SequenceEqual(other.Cells);

    public override bool Equals(object? other)
        => other is Row row && Equals(row);

    public override int GetHashCode()
        => Cells.GetHashCode();
}
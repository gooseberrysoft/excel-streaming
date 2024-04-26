using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Tests.Excel;

[StructLayout(LayoutKind.Auto)]
public readonly struct Row : IEquatable<Row>
{
    public Row(IReadOnlyCollection<Cell> cells, decimal? height = null)
    {
        Cells = cells;
        Height = height;
    }

    public decimal? Height { get; }

    public IReadOnlyCollection<Cell> Cells { get; }

    public bool Equals(Row other)
        => Cells.SequenceEqual(other.Cells);

    public override bool Equals(object? other)
        => other is Row row && Equals(row);

    public override int GetHashCode()
        => Cells.GetHashCode();
}
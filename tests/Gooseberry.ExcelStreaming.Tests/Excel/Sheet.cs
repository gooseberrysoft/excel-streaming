using System.Runtime.InteropServices;
using Gooseberry.ExcelStreaming.Tests.Extensions;

namespace Gooseberry.ExcelStreaming.Tests.Excel;

[StructLayout(LayoutKind.Auto)]
public readonly record struct Sheet(
    string Name,
    IReadOnlyCollection<Row> Rows,
    IReadOnlyCollection<Column>? Columns = null,
    IReadOnlyCollection<string>? Merges = null,
    IReadOnlyCollection<Picture>? Pictures = null)
{
    public IReadOnlyCollection<Picture> Pictures { get; init; } = Pictures ?? Array.Empty<Picture>();

    public IReadOnlyCollection<Column> Columns { get; init; } = Columns ?? Array.Empty<Column>();

    public IReadOnlyCollection<string> Merges { get; init; } = Merges ?? Array.Empty<string>();

    public bool Equals(Sheet other)
    {
        return string.Equals(Name, other.Name, StringComparison.Ordinal)
            && Rows.SequenceEqual(other.Rows)
            && Merges.SequenceEqual(other.Merges)
            && Columns.SequenceEqual(other.Columns)
            && Pictures.SequenceEqual(other.Pictures);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(
            Name,
            Rows.GetCollectionHashCode(),
            Merges.GetCollectionHashCode(),
            Columns.GetCollectionHashCode(),
            Pictures.GetCollectionHashCode());
    }
}
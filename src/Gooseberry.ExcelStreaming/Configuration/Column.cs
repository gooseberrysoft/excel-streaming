// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

public sealed record Column(
    Column.Range? ColumnsRange = null,
    decimal? Width = null,
    bool IsHidden = false)
{
    public Column(decimal width) : this(Width: width)
    {
    }

    public readonly struct Range
    {
        public int MinIndex { get; }
        public int MaxIndex { get; }

        public Range(int index) : this(index, index)
        {
        }

        public Range(int minIndex, int maxIndex)
        {
            if (minIndex < 1)
                throw new ArgumentOutOfRangeException(nameof(minIndex), minIndex,
                    $"{nameof(minIndex)} should be greater than 0. Index of first column is 1.");

            if (maxIndex < minIndex)
                throw new ArgumentException($"{nameof(minIndex)} should be less or equal to {nameof(maxIndex)}.");

            MinIndex = minIndex;
            MaxIndex = maxIndex;
        }
    }
}
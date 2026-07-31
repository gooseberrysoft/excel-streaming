using System.Buffers;
using System.Collections;
using System.Diagnostics.CodeAnalysis;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class SequenceList<T> : IDisposable, IEnumerable<ReadOnlyMemory<T>>
{
    private const int q = 2;
    private const int b1 = 256;
    private static readonly double LnQ = Math.Log(q);

    private int length;
    private int lastIndex = -1;
    private readonly List<T[]> buckets = [];


    /// <summary>
    /// Not thread-safe
    /// </summary>
    /// <returns></returns>
    public int Next()
    {
        lastIndex++;

        if (lastIndex >= length)
            Extend();

        return lastIndex;
    }

    public int Length => length;

    /// <summary>
    /// Thread-safe
    /// </summary>
    /// <param name="index"></param>
    /// <returns></returns>
    public T this[int index]
    {
        set => SetValue(index, value);
        get => GetValue(index);
    }

    public IEnumerator<ReadOnlyMemory<T>> GetEnumerator()
    {
        var bucketNumber = 1;
        foreach (var bucket in buckets)
        {
            var count = (int)(b1 * Math.Pow(q, bucketNumber));

            yield return bucket.AsMemory(0, count);

            bucketNumber++;
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public void Dispose()
    {
        foreach (var bucket in buckets)
            ArrayPool<T>.Shared.Return(bucket, clearArray: true);

        buckets.Clear();

        length = 0;
    }

    private void Extend()
    {
        var nextSize = (int)(b1 * Math.Pow(q, buckets.Count));
        length += nextSize;

        var rentedArray = ArrayPool<T>.Shared.Rent(nextSize);
        buckets.Add(rentedArray);
    }

    private void SetValue(int index, T value)
    {
        if (index > lastIndex)
            ThrowInvalidArgument(index);

        var bucketNumber = GetBucketNumber(index);
        var bucketElementIndex = GetBucketElementIndex(bucketNumber, index);

        buckets[bucketNumber - 1][bucketElementIndex] = value;
    }

    private T GetValue(int index)
    {
        if (index > lastIndex)
            ThrowInvalidArgument(index);

        var bucketNumber = GetBucketNumber(index);
        var bucketElementIndex = GetBucketElementIndex(bucketNumber, index);

        return buckets[bucketNumber - 1][bucketElementIndex];
    }

    private static int GetBucketNumber(int elementIndex)
    {
        return elementIndex switch
        {
            < b1 => 1,
            < b1 * (q * q - 1) / (q - 1) => 2,
            < b1 * (q * q * q - 1) / (q - 1) => 3,
            < b1 * (q * q * q * q - 1) / (q - 1) => 4,
            < b1 * (q * q * q * q * q - 1) / (q - 1) => 5,
            < b1 * (q * q * q * q * q * q - 1) / (q - 1) => 6,
            < b1 * (q * q * q * q * q * q * q - 1) / (q - 1) => 7,
            < b1 * (q * q * q * q * q * q * q * q - 1) / (q - 1) => 8,

            _ => (int)Math.Floor(Math.Log((elementIndex * (q - 1.0)) / b1 + 1.0) / LnQ)
        };
    }

    private static int GetBucketElementIndex(int bucketNumber, int elementIndex)
    {
        var offset = bucketNumber switch
        {
            1 => 0,
            2 => b1,
            3 => b1 * (q * q - 1) / (q - 1),
            4 => b1 * (q * q * q - 1) / (q - 1),
            5 => b1 * (q * q * q * q - 1) / (q - 1),
            6 => b1 * (q * q * q * q * q - 1) / (q - 1),
            7 => b1 * (q * q * q * q * q * q - 1) / (q - 1),
            8 => b1 * (q * q * q * q * q * q * q - 1) / (q - 1),

            _ => (b1 * (Math.Pow(q, bucketNumber - 1) - 1.0)) / (q - 1.0)
        };

        return (int)(elementIndex - offset);
    }

    [DoesNotReturn]
    private void ThrowInvalidArgument(int index)
        => throw new ArgumentException($"Index must be less than {lastIndex}, index: {index}");
}
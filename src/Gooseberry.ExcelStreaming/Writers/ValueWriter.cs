using System;

namespace Gooseberry.ExcelStreaming.Writers;

internal abstract class ValueWriter<T> : IValueWriter<T>
{
    protected abstract int MaximumChars { get; }
    protected abstract bool TryFormat(in T value, Span<byte> destination, out int encodedBytes);

    public void WriteValue(in T value, BuffersChain bufferWriter, ref Span<byte> destination, ref int written)
    {
        if (TryFormat(value, destination, out var encodedBytes))
        {
            destination = destination.Slice(encodedBytes);
            written += encodedBytes;
            return;
        }

        bufferWriter.Advance(written);
        destination = bufferWriter.GetSpan(MaximumChars);
        written = 0;

        if (!TryFormat(value, destination, out encodedBytes))
            throw new InvalidOperationException($"Can't format {typeof(T)}. Not enough memory");

        destination = destination.Slice(encodedBytes);
        written += encodedBytes;
    }
}
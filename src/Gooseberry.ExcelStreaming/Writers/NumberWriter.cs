using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class NumberWriter<T, TFormatter>
    where TFormatter : INumberFormatter<T>
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteValue(in T value, BuffersChain bufferWriter, ref Span<byte> destination, ref int written)
    {
        if (TFormatter.TryFormat(value, destination, out var encodedBytes))
        {
            destination = destination.Slice(encodedBytes);
            written += encodedBytes;
            return;
        }

        WriteAdvance(value, bufferWriter, ref destination, ref written);
    }

    private static void WriteAdvance(in T value, BuffersChain bufferWriter, ref Span<byte> destination, ref int written)
    {
        bufferWriter.Advance(written);
        destination = bufferWriter.GetSpan(TFormatter.MaximumChars);
        written = 0;

        var encodedBytes = 0;
        var attempt = 1;

        while (!TFormatter.TryFormat(value, destination, out encodedBytes))
        {
            if (attempt > 10)
                throw new InvalidOperationException($"Can't format {typeof(T)}.");

            attempt++;
            destination = bufferWriter.GetSpan(TFormatter.MaximumChars * attempt);
        }

        destination = destination.Slice(encodedBytes);
        written += encodedBytes;
    }
}
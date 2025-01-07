#if NET8_0_OR_GREATER
using System.Globalization;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class Utf8SpanFormattableWriter
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteValue<T>(
        in T value,
        ReadOnlySpan<char> format,
        IFormatProvider? provider,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written) where T : IUtf8SpanFormattable
    {
        if (value.TryFormat(destination, out var bytesWritten, format, provider ?? CultureInfo.InvariantCulture))
        {
            destination = destination.Slice(bytesWritten);
            written += bytesWritten;
            return;
        }

        WriteAdvance(value, format, provider, buffer, ref destination, ref written);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteValue<T>(
        in T value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written) where T : IUtf8SpanFormattable
    {
        WriteValue(value, format: ReadOnlySpan<char>.Empty, null, buffer, ref destination, ref written);
    }

    private static void WriteAdvance<T>(
        in T value,
        ReadOnlySpan<char> format,
        IFormatProvider? provider,
        BuffersChain bufferWriter,
        ref Span<byte> destination,
        ref int written) where T : IUtf8SpanFormattable
    {
        bufferWriter.Advance(written);
        destination = bufferWriter.GetSpan(Math.Max(destination.Length * 2, Buffer.MinSize));
        written = 0;

        var bytesWritten = 0;
        var attempt = 1;

        while (!value.TryFormat(destination, out bytesWritten, format, provider ?? CultureInfo.InvariantCulture))
        {
            if (attempt > 10)
                throw new InvalidOperationException($"Can't format {typeof(T)}.");

            attempt++;
            destination = bufferWriter.GetSpan(destination.Length * 2);
        }

        destination = destination.Slice(bytesWritten);
        written += bytesWritten;
    }
}

#endif
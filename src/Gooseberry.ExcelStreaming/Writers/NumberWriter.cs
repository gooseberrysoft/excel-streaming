using System;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class NumberWriter
{
    private static readonly IntFormatter IntFormatter = new();
    private static readonly LongFormatter LongFormatter = new();
    private static readonly DecimalFormatter DecimalFormatter = new();

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this int value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written) 
        => WriteTo(value, IntFormatter, buffer, ref destination, ref written);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this uint value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written) 
        => WriteTo(value, LongFormatter, buffer, ref destination, ref written);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this long value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written) 
        => WriteTo(value, LongFormatter, buffer, ref destination, ref written);
    
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this decimal value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written) 
        => WriteTo(value, DecimalFormatter, buffer, ref destination, ref written);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo<T>(
        this T value, 
        INumberFormatter<T> formatter,
        BuffersChain buffer, 
        ref Span<byte> destination, 
        ref int written)
    {
        if (formatter.TryFormat(value, destination, out var encodedBytes))
        {
            destination = destination.Slice(encodedBytes);
            written += encodedBytes;
            return;
        }

        buffer.Advance(written);
        destination = buffer.GetSpan(formatter.MaximumChars);
        written = 0;

        if (!formatter.TryFormat(value, destination, out encodedBytes))
            throw new InvalidOperationException($"Can't format {typeof(T)}. Not enough memory");

        destination = destination.Slice(encodedBytes);
        written += encodedBytes;
    }
}
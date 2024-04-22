using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class NumberWriterExtensions
{
    private static readonly NumberWriter<int, IntFormatter> IntWriter = new();
    private static readonly NumberWriter<long, LongFormatter> LongWriter = new();
    private static readonly NumberWriter<decimal, DecimalFormatter> DecimalWriter = new();

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this int value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
        => IntWriter.WriteValue(value, buffer, ref destination, ref written);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this long value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
        => LongWriter.WriteValue(value, buffer, ref destination, ref written);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this uint value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written) 
        => LongWriter.WriteValue(value, buffer, ref destination, ref written);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this decimal value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written) 
        => DecimalWriter.WriteValue(value, buffer, ref destination, ref written);
}

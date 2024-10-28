using Gooseberry.ExcelStreaming.Writers.Formatters;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class NumberWriterExtensions
{
    internal static void WriteTo(
        this int value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
        => NumberWriter<int, IntFormatter>.WriteValue(value, buffer, ref destination, ref written);

    internal static void WriteHex8To(
        this int value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
        => NumberWriter<int, IntHex8Formatter>.WriteValue(value, buffer, ref destination, ref written);

    internal static void WriteTo(
        this long value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
        => NumberWriter<long, LongFormatter>.WriteValue(value, buffer, ref destination, ref written);

    internal static void WriteTo(
        this uint value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
        => NumberWriter<uint, UIntFormatter>.WriteValue(value, buffer, ref destination, ref written);

    internal static void WriteTo(
        this decimal value,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
        => NumberWriter<decimal, DecimalFormatter>.WriteValue(value, buffer, ref destination, ref written);
}
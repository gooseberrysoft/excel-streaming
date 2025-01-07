#if NET8_0_OR_GREATER
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class Utf8DateTimeCellWriter
{
    private const int NumberSize = 20;

    private static ReadOnlySpan<byte> StylePrefix => "<c s=\""u8;
    private static ReadOnlySpan<byte> StylePostfix => "\"><v>"u8;
    private static ReadOnlySpan<byte> Postfix => "</v></c>"u8;

    private static readonly int Size = StylePrefix.Length + NumberSize + StylePostfix.Length
        + NumberSize
        + Postfix.Length;

    public static void Write(
        DateTime value,
        ReadOnlySpan<char> format,
        IFormatProvider? provider,
        BuffersChain buffer,
        StyleReference style)
    {
        var span = buffer.GetSpan(Size);
        var written = 0;

        StylePrefix.CopyTo(ref span, ref written);
        Utf8SpanFormattableWriter.WriteValue(style.Value, buffer, ref span, ref written);
        StylePostfix.CopyTo(ref span, ref written);

        Utf8SpanFormattableWriter.WriteValue(value.ToOADate(), format, provider, buffer, ref span, ref written);

        Postfix.CopyTo(span);
        buffer.Advance(written + Postfix.Length);
    }
}
#endif
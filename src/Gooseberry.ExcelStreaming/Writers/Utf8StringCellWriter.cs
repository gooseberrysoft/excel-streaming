#if NET8_0_OR_GREATER
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class Utf8StringCellWriter
{
    private static ReadOnlySpan<byte> Prefix => "<c t=\"str\"><v>"u8;
    private static ReadOnlySpan<byte> StylePrefix => "<c t=\"str\" s=\""u8;
    private static ReadOnlySpan<byte> StylePostfix => "\"><v>"u8;
    private static ReadOnlySpan<byte> Postfix => "</v></c>"u8;

    private const int NumberSize = 20;
    private static readonly int Size = Prefix.Length + Postfix.Length;

    private static readonly int StyleSize = StylePrefix.Length + NumberSize + StylePostfix.Length
        + Postfix.Length;

    public static void Write<T>(
        T value,
        ReadOnlySpan<char> format,
        IFormatProvider? provider,
        BuffersChain buffer)
        where T : IUtf8SpanFormattable
    {
        var span = buffer.GetSpan(Size);
        var written = 0;

        Prefix.CopyTo(ref span, ref written);

        StringWriter.WriteEscapedUtf8To(value, format, provider, buffer, ref span, ref written);

        Postfix.WriteAdvanceTo(buffer, span, written);
    }

    public static void Write<T>(
        T value,
        ReadOnlySpan<char> format,
        IFormatProvider? provider,
        BuffersChain buffer,
        StyleReference style)
        where T : IUtf8SpanFormattable
    {
        var span = buffer.GetSpan(StyleSize);
        var written = 0;

        StylePrefix.CopyTo(ref span, ref written);
        Utf8SpanFormattableWriter.WriteValue(style.Value, buffer, ref span, ref written);
        StylePostfix.CopyTo(ref span, ref written);
        
        StringWriter.WriteEscapedUtf8To(value, format, provider, buffer, ref span, ref written);
        Postfix.WriteAdvanceTo(buffer, span, written);
    }
}
#endif
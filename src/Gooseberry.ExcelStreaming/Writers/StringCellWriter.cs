using System.Text;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class StringCellWriter
{
    // https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3?ui=en-us&rs=en-us&ad=us#ID0EBABAAA=Excel_2016-2013
    private const int MaxCharacters = 32_767;
    internal const int MaxBytes = MaxCharacters * 3;

    private const int MaxValueSizeBuffer = 1024;
    private static ReadOnlySpan<byte> Prefix => "<c t=\"str\"><v>"u8;
    private static ReadOnlySpan<byte> Postfix => "</v></c>"u8;
    private static readonly int Size = Prefix.Length + Postfix.Length;

    private static ReadOnlySpan<byte> StylePrefix => "<c t=\"str\" s=\""u8;
    private static ReadOnlySpan<byte> StylePostfix => "\"><v>"u8;
    private const int NumberSize = 20;

    private static readonly int StyleSize = StylePrefix.Length + NumberSize + StylePostfix.Length
        + Postfix.Length;


    public static void Write(ReadOnlySpan<char> value, BuffersChain buffer, Encoder encoder)
    {
        if (value.Length > MaxCharacters)
            ThrowCharsLimitExceeded();

        var span = buffer.GetSpan(Size + Math.Min(value.Length, MaxValueSizeBuffer));
        var written = 0;

        Prefix.CopyTo(ref span, ref written);
        value.WriteEscapedTo(buffer, encoder, ref span, ref written);

        Postfix.WriteAdvanceTo(buffer, span, written);
    }

    public static void Write(ReadOnlySpan<char> value, BuffersChain buffer, Encoder encoder, StyleReference style)
    {
        if (value.Length > MaxCharacters)
            ThrowCharsLimitExceeded();

        var span = buffer.GetSpan(StyleSize + Math.Min(value.Length, MaxValueSizeBuffer));
        var written = 0;

        StylePrefix.CopyTo(ref span, ref written);
#if NET8_0_OR_GREATER
        Utf8SpanFormattableWriter.WriteValue(style.Value, buffer, ref span, ref written);
#else
        style.Value.WriteTo(buffer, ref span, ref written);
#endif
        StylePostfix.CopyTo(ref span, ref written);

        value.WriteEscapedTo(buffer, encoder, ref span, ref written);

        Postfix.WriteAdvanceTo(buffer, span, written);
    }

    public static void WriteUtf8(ReadOnlySpan<byte> value, BuffersChain buffer)
    {
        if (value.Length > MaxBytes)
            ThrowCharsLimitExceeded();

        var span = buffer.GetSpan(Prefix.Length + Postfix.Length + value.Length);
        var written = 0;

        Prefix.CopyTo(ref span, ref written);
        value.WriteEscapedUtf8To(buffer, ref span, ref written);

        Postfix.WriteAdvanceTo(buffer, span, written);
    }

    public static void WriteUtf8(ReadOnlySpan<byte> value, BuffersChain buffer, StyleReference style)
    {
        if (value.Length > MaxBytes)
            ThrowCharsLimitExceeded();
        var span = buffer.GetSpan(StyleSize + Math.Min(value.Length, MaxValueSizeBuffer));
        var written = 0;

        StylePrefix.CopyTo(ref span, ref written);
#if NET8_0_OR_GREATER
        Utf8SpanFormattableWriter.WriteValue(style.Value, buffer, ref span, ref written);
#else
        style.Value.WriteTo(buffer, ref span, ref written);
#endif
        StylePostfix.CopyTo(ref span, ref written);

        value.WriteEscapedUtf8To(buffer, ref span, ref written);

        Postfix.WriteAdvanceTo(buffer, span, written);
    }

    public static void ThrowCharsLimitExceeded()
        => throw new ArgumentException($"Cell value exceed Excel {MaxCharacters} chars limit.");
}
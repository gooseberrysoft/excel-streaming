using System.Runtime.CompilerServices;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers.Cells;

internal static class StringCellWriter
{
    // https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3?ui=en-us&rs=en-us&ad=us#ID0EBABAAA=Excel_2016-2013
    private const int MaxCharacters = 32_767;
    internal const int MaxBytes = MaxCharacters * 3;

    private static ReadOnlySpan<byte> Prefix => "<c t=\"str\"><v>"u8;
    private static ReadOnlySpan<byte> ClosedPrefix => "</v></c><c t=\"str\"><v>"u8;

    private static ReadOnlySpan<byte> StylePrefix => "<c t=\"str\" s=\""u8;
    private static ReadOnlySpan<byte> ClosedStylePrefix => "</v></c><c t=\"str\" s=\""u8;
    private static ReadOnlySpan<byte> StylePostfix => "\"><v>"u8;
    private static int DefaultBufferSize = ClosedStylePrefix.Length + 25;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void Write(ReadOnlySpan<char> value, CellWritingContext context, StyleReference? style)
    {
        if (value.Length > MaxCharacters)
            ThrowCharsLimitExceeded();

        var buffer = context.Buffer;
        var span = buffer.GetSpan(DefaultBufferSize);
        var written = 0;

        if (style.HasValue)
        {
            (context.IsCellValueOpened ? ClosedStylePrefix : StylePrefix).CopyTo(ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(style.Value.Value, buffer, ref span, ref written);
            StylePostfix.CopyTo(ref span, ref written);
        }
        else
            (context.IsCellValueOpened ? ClosedPrefix : Prefix).CopyTo(ref span, ref written);

        value.WriteEscapedTo(buffer, context.Encoder, ref span, ref written);

        buffer.Advance(written);
        context.OpenCellValue();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void WriteUtf8(ReadOnlySpan<byte> value, CellWritingContext context, StyleReference? style)
    {
        if (value.Length > MaxBytes)
            ThrowCharsLimitExceeded();

        var buffer = context.Buffer;
        var span = buffer.GetSpan(DefaultBufferSize);
        var written = 0;

        if (style.HasValue)
        {
            (context.IsCellValueOpened ? ClosedStylePrefix : StylePrefix).CopyTo(ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(style.Value.Value, buffer, ref span, ref written);
            StylePostfix.CopyTo(ref span, ref written);
        }
        else
            (context.IsCellValueOpened ? ClosedPrefix : Prefix).CopyTo(ref span, ref written);

        value.WriteEscapedUtf8To(buffer, ref span, ref written);

        buffer.Advance(written);
        context.OpenCellValue();
    }

    public static void ThrowCharsLimitExceeded()
        => throw new ArgumentException($"Cell value exceed Excel {MaxCharacters} chars limit.");
}
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class EmptyCellWriter
{
    private static ReadOnlySpan<byte> EmptyCell => "<c></c>"u8;
    private static ReadOnlySpan<byte> StylePrefix => "<c s=\""u8;
    private static ReadOnlySpan<byte> StylePostfix => "\"></c>"u8;

    public static void Write(BufferSequence buffer, StyleReference? style = null)
    {
        var span = buffer.GetSpan(EmptyCell.Length);
        var written = 0;

        if (style.HasValue)
        {
            StylePrefix.WriteTo(buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(style.Value.Value, buffer, ref span, ref written);
            StylePostfix.WriteTo(buffer, ref span, ref written);
            
            buffer.Advance(written);
            return;
        }

        EmptyCell.WriteAdvanceTo(buffer, span, written);
    }
}
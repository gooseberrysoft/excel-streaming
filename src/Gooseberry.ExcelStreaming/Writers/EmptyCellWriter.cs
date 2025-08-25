using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers.Formatters;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class EmptyCellWriter
{
    private static ReadOnlySpan<byte> EmptyCell => "<c></c>"u8;
    private static ReadOnlySpan<byte> StylePrefix => "<c s=\""u8;
    private static ReadOnlySpan<byte> StylePostfix => "\"></c>"u8;

    public static void Write(BuffersChain buffer, StyleReference? style = null)
    {
        var span = buffer.GetSpan(EmptyCell.Length);
        var written = 0;

        if (style.HasValue)
        {
            StylePrefix.WriteTo(buffer, ref span, ref written);
#if NET8_0_OR_GREATER
            Utf8SpanFormattableWriter.WriteValue(style.Value.Value, buffer, ref span, ref written);
#else
            style.Value.Value.WriteTo(buffer, ref span, ref written);
#endif
            StylePostfix.WriteTo(buffer, ref span, ref written);
            
            buffer.Advance(written);
            return;
        }

        EmptyCell.WriteAdvanceTo(buffer, span, written);
    }
}
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers.Cells;

internal static class EmptyCellWriter
{
    private static ReadOnlySpan<byte> EmptyCell => "<c></c>"u8;
    private static ReadOnlySpan<byte> ClosedEmptyCell => "</v></c><c></c>"u8;
    private static ReadOnlySpan<byte> ClosedStylePrefix => "</v></c><c s=\""u8;
    private static ReadOnlySpan<byte> StylePrefix => "<c s=\""u8;
    private static ReadOnlySpan<byte> StylePostfix => "\"></c>"u8;

    public static void Write(CellWritingContext context, StyleReference? style = null)
    {
        var buffer = context.Buffer;
        var span = buffer.GetSpan(ClosedEmptyCell.Length);
        var written = 0;

        if (style.HasValue)
        {
            (context.IsCellValueOpened ? ClosedStylePrefix : StylePrefix).WriteTo(buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(style.Value.Value, buffer, ref span, ref written);
            StylePostfix.WriteTo(buffer, ref span, ref written);

            buffer.Advance(written);
            context.CloseCellValue();
            return;
        }

        (context.IsCellValueOpened ? ClosedEmptyCell : EmptyCell).WriteAdvanceTo(buffer, span, written);

        context.CloseCellValue();
    }
}
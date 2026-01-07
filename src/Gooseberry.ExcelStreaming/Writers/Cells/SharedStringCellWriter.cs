using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers.Cells;

internal static class SharedStringCellWriter
{
    public static void Write(SharedStringReference sharedString, CellWritingContext context, StyleReference? style = null)
    {
        var buffer = context.Buffer;
        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            (context.IsCellValueOpened ? "</v></c><c t=\"s\" s=\""u8 : "<c t=\"s\" s=\""u8).WriteTo(buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(style.Value.Value, buffer, ref span, ref written);
            "\"><v>"u8.WriteTo(buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(sharedString.Value, buffer, ref span, ref written);
        }
        else
        {
            (context.IsCellValueOpened ? "</v></c><c t=\"s\"><v>"u8 : "<c t=\"s\"><v>"u8).WriteTo(buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(sharedString.Value, buffer, ref span, ref written);
        }

        context.OpenCellValue();

        //"</v></c>"u8.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
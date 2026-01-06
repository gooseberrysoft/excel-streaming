using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class SharedStringCellWriter
{
    public static void Write(SharedStringReference sharedString, BuffersChain buffer, StyleReference? style = null)
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            "<c t=\"s\" s=\""u8.WriteTo(buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(style.Value.Value, buffer, ref span, ref written);
            "\"><v>"u8.WriteTo(buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(sharedString.Value, buffer, ref span, ref written);
        }
        else
        {
            "<c t=\"s\"><v>"u8.WriteTo(buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(sharedString.Value, buffer, ref span, ref written);
        }

        "</v></c>"u8.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
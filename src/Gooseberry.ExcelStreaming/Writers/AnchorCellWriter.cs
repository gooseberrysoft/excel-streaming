namespace Gooseberry.ExcelStreaming.Writers;

internal static class AnchorCellWriter
{
    public static void Write(in AnchorCell cell, BufferSequence buffer)
    {
        var written = 0;
        var span = buffer.GetSpan();

        Write(cell, buffer, ref span, ref written);
        buffer.Advance(written);
    }

    public static void Write(in AnchorCell cell, BufferSequence buffer, ref Span<byte> span, ref int written)
    {
        "<xdr:col>"u8.WriteTo(buffer, ref span, ref written);
        Utf8SpanFormattableWriter.WriteValue(cell.Column, buffer, ref span, ref written);
        "</xdr:col>"u8.WriteTo(buffer, ref span, ref written);

        "<xdr:colOff>"u8.WriteTo(buffer, ref span, ref written);
        Utf8SpanFormattableWriter.WriteValue(cell.Offset.X, buffer, ref span, ref written);
        "</xdr:colOff>"u8.WriteTo(buffer, ref span, ref written);

        "<xdr:row>"u8.WriteTo(buffer, ref span, ref written);
        Utf8SpanFormattableWriter.WriteValue(cell.Row, buffer, ref span, ref written);
        "</xdr:row>"u8.WriteTo(buffer, ref span, ref written);

        "<xdr:rowOff>"u8.WriteTo(buffer, ref span, ref written);
        Utf8SpanFormattableWriter.WriteValue(cell.Offset.Y, buffer, ref span, ref written);
        "</xdr:rowOff>"u8.WriteTo(buffer, ref span, ref written);
    }
}
namespace Gooseberry.ExcelStreaming.Writers;

internal static class RowWriter
{
    public static void WriteStartRow(CellWritingContext context, bool rowStarted, in RowAttributes rowAttributes)
    {
        var buffer = context.Buffer;
        var span = buffer.GetSpan();
        var written = 0;
        var attributeIsEmpty = rowAttributes.IsEmpty();

        if (rowStarted && attributeIsEmpty)
        {
            (context.IsCellValueOpened ? "</v></c></row><row>"u8 : "</row><row>"u8).WriteTo(buffer, ref span, ref written);
            buffer.Advance(written);

            context.CloseCellValue();
            return;
        }

        if (rowStarted)
            (context.IsCellValueOpened ? "</v></c></row><row"u8 : "</row><row"u8).WriteTo(buffer, ref span, ref written);
        else
            "<row"u8.WriteTo(buffer, ref span, ref written);

        context.CloseCellValue();

        if (!attributeIsEmpty)
            AddAttributes(buffer, ref span, ref written, rowAttributes);

        ">"u8.WriteTo(buffer, ref span, ref written);
        buffer.Advance(written);
    }

    public static void WriteEndRow(CellWritingContext context)
    {
        var buffer = context.Buffer;
        var span = buffer.GetSpan();
        var written = 0;

        (context.IsCellValueOpened ? "</v></c></row>"u8 : "</row>"u8).WriteTo(buffer, ref span, ref written);
        
        context.CloseCellValue();
        buffer.Advance(written);
    }

    private static void AddAttributes(
        BuffersChain buffer,
        ref Span<byte> span,
        ref int written,
        in RowAttributes rowAttributes)
    {
        if (rowAttributes.Height.HasValue)
        {
            " ht=\""u8.WriteTo(buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(rowAttributes.Height.Value, buffer, ref span, ref written);
            "\" customHeight=\"1\""u8.WriteTo(buffer, ref span, ref written);
        }

        if (rowAttributes.OutlineLevel.HasValue)
        {
            " outlineLevel=\""u8.WriteTo(buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(rowAttributes.OutlineLevel.Value, buffer, ref span, ref written);
            "\""u8.WriteTo(buffer, ref span, ref written);
        }

        if (rowAttributes.IsHidden.HasValue && rowAttributes.IsHidden.Value)
            " hidden=\"true\""u8.WriteTo(buffer, ref span, ref written);

        if (rowAttributes.IsCollapsed.HasValue && rowAttributes.IsCollapsed.Value)
            " collapsed=\"1\""u8.WriteTo(buffer, ref span, ref written);
    }
}
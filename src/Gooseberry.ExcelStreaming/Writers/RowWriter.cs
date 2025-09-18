namespace Gooseberry.ExcelStreaming.Writers;

internal static class RowWriter
{
    public static void WriteStartRow(BuffersChain buffer, bool rowStarted, in RowAttributes rowAttributes)
    {
        var span = buffer.GetSpan();
        var written = 0;
        var attributeIsEmpty = rowAttributes.IsEmpty();

        if (rowStarted && attributeIsEmpty)
        {
            "</row><row>"u8.WriteTo(buffer, ref span, ref written);
            buffer.Advance(written);

            return;
        }

        if (rowStarted)
            "</row>"u8.WriteTo(buffer, ref span, ref written);

        "<row"u8.WriteTo(buffer, ref span, ref written);

        if (!attributeIsEmpty)
            AddAttributes(buffer, ref span, ref written, rowAttributes);

        ">"u8.WriteTo(buffer, ref span, ref written);
        buffer.Advance(written);
    }

    public static void WriteEndRow(BuffersChain buffer)
    {
        var span = buffer.GetSpan();
        var written = 0;

        "</row>"u8.WriteTo(buffer, ref span, ref written);

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
            rowAttributes.Height.Value.WriteTo(buffer, ref span, ref written);
            "\" customHeight=\"1\""u8.WriteTo(buffer, ref span, ref written);
        }

        if (rowAttributes.OutlineLevel.HasValue)
        {
            " outlineLevel=\""u8.WriteTo(buffer, ref span, ref written);
            NumberWriterExtensions.WriteTo(rowAttributes.OutlineLevel.Value, buffer, ref span, ref written);
            "\""u8.WriteTo(buffer, ref span, ref written);
        }

        if (rowAttributes.IsHidden.HasValue && rowAttributes.IsHidden.Value)
            " hidden=\"true\""u8.WriteTo(buffer, ref span, ref written);

        if (rowAttributes.IsCollapsed.HasValue && rowAttributes.IsCollapsed.Value)
            " collapsed=\"1\""u8.WriteTo(buffer, ref span, ref written);
    }
}
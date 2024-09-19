using Gooseberry.ExcelStreaming.Extensions;

namespace Gooseberry.ExcelStreaming.Writers;

using SheetDataRow = Constants.Worksheet.SheetData.Row;

internal static class RowWriter
{
    private static readonly byte[] RowCloseAndStartWithoutAttributes = SheetDataRow.Postfix
            .Combine(SheetDataRow.Open.Prefix, SheetDataRow.Open.Postfix);

    public static void WriteStartRow(BuffersChain buffer, bool rowStarted, in RowAttributes rowAttributes)
    {
        var span = buffer.GetSpan();
        var written = 0;
        var attributeIsEmpty = rowAttributes.IsEmpty();
        
        if (rowStarted && attributeIsEmpty)
        {
            RowCloseAndStartWithoutAttributes.WriteTo(buffer, ref span, ref written);
            buffer.Advance(written);

            return;
        }

        if (rowStarted)
            SheetDataRow.Postfix.WriteTo(buffer, ref span, ref written);

        SheetDataRow.Open.Prefix.WriteTo(buffer, ref span, ref written);

        if (!attributeIsEmpty)
            AddAttributes(buffer, ref span, ref written, rowAttributes);

        SheetDataRow.Open.Postfix.WriteTo(buffer, ref span, ref written);
        buffer.Advance(written);
    }

    public static void WriteEndRow(BuffersChain buffer)
    {
        var span = buffer.GetSpan();
        var written = 0;

        SheetDataRow.Postfix.WriteTo(buffer, ref span, ref written);

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
            SheetDataRow.Open.Height.Prefix.WriteTo(buffer, ref span, ref written);
            rowAttributes.Height.Value.WriteTo(buffer, ref span, ref written);
            SheetDataRow.Open.Height.Postfix.WriteTo(buffer, ref span, ref written);
        }

        if (rowAttributes.OutlineLevel.HasValue)
        {
            SheetDataRow.Open.OutlineLevel.Prefix.WriteTo(buffer, ref span, ref written);
            NumberWriterExtensions.WriteTo(rowAttributes.OutlineLevel.Value, buffer, ref span, ref written);
            SheetDataRow.Open.OutlineLevel.Postfix.WriteTo(buffer, ref span, ref written);
        }

        if (rowAttributes.IsHidden.HasValue && rowAttributes.IsHidden.Value)
            SheetDataRow.Open.Hidden.WriteTo(buffer, ref span, ref written);

        if (rowAttributes.IsCollapsed.HasValue && rowAttributes.IsCollapsed.Value)
            SheetDataRow.Open.Collapsed.WriteTo(buffer, ref span, ref written);
    }
}
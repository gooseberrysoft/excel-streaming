using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class SheetWriter
{
    public static void WriteStartSheet(BuffersChain buffer, in SheetConfiguration? configuration)
    {
        var span = buffer.GetSpan();
        var written = 0;

        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"u8
            .WriteTo(buffer, ref span, ref written);

        if (configuration != null)
            WriteSheetView(configuration, buffer, ref span, ref written);

        if (configuration?.Columns != null && configuration?.Columns.Count != 0)
            WriteColumns(configuration!.Columns, buffer, ref span, ref written);

        "<sheetData>"u8.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }

    public static void WriteEndSheet(
        BuffersChain buffer,
        Encoder encoder,
        Drawing drawing,
        IReadOnlyCollection<Merge> merges,
        in CellRange? autoFilter,
        Dictionary<string, List<CellReference>>? hyperlinks)
    {
        var span = buffer.GetSpan();
        var written = 0;

        "</sheetData>"u8.WriteTo(buffer, ref span, ref written);

        if (autoFilter != null)
            WriteAutoFilter(autoFilter.Value, buffer, encoder, ref span, ref written);

        WriteMerges(merges, buffer, ref span, ref written);
        
        if (hyperlinks != null)
            WriteHyperlinks(hyperlinks, buffer, ref span, ref written);

        WriteDrawing(drawing, buffer, encoder, ref span, ref written);

        "</worksheet>"u8.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }

    private static void WriteSheetView(
        in SheetConfiguration configuration,
        BuffersChain buffer,
        ref Span<byte> span,
        ref int written)
    {
        "<sheetViews><sheetView workbookViewId=\"0\""u8.WriteTo(buffer, ref span, ref written);

        WriteShowGridLines(configuration.ShowGridLines, buffer, ref span, ref written);

        ">"u8.WriteTo(buffer, ref span, ref written);

        if (configuration.FrozenColumns.HasValue || configuration.FrozenRows.HasValue)
        {
            WriteFreezePanes(
                new CellReference(
                    (configuration.FrozenColumns ?? 0) + 1,
                    (configuration.FrozenRows ?? 0) + 1),
                buffer, ref span, ref written);
        }

        "</sheetView></sheetViews>"u8.WriteTo(buffer, ref span, ref written);
    }

    private static void WriteFreezePanes(
        CellReference cellReference,
        BuffersChain buffer,
        ref Span<byte> span,
        ref int written)
    {
        "<pane topLeftCell=\""u8.WriteTo(buffer, ref span, ref written);
        cellReference.WriteTo(buffer, ref span, ref written);

        "\" ySplit=\""u8.WriteTo(buffer, ref span, ref written);
        (cellReference.Row - 1).WriteTo(buffer, ref span, ref written);

        "\" xSplit=\""u8.WriteTo(buffer, ref span, ref written);
        (cellReference.Column - 1).WriteTo(buffer, ref span, ref written);

        "\" activePane=\"bottomRight\" state=\"frozen\"/>"u8.WriteTo(buffer, ref span, ref written);
    }

    private static void WriteShowGridLines(
        bool showGridLines,
        BuffersChain buffer,
        ref Span<byte> span,
        ref int written)
    {
        " showGridLines=\""u8.WriteTo(buffer, ref span, ref written);
        showGridLines.WriteTo(buffer, ref span, ref written);
        "\""u8.WriteTo(buffer, ref span, ref written);
    }

    private static void WriteAutoFilter(
        in CellRange filterRange,
        BuffersChain buffer,
        Encoder encoder,
        ref Span<byte> span,
        ref int written)
    {
        "<autoFilter ref=\""u8.WriteTo(buffer, ref span, ref written);

        if (!string.IsNullOrWhiteSpace(filterRange.Range))
        {
            filterRange.Range.WriteTo(buffer, encoder, ref span, ref written);
        }
        else
        {
            filterRange.FromCell.WriteTo(buffer, ref span, ref written);
            ":"u8.WriteTo(buffer, ref span, ref written);
            filterRange.ToCell.WriteTo(buffer, ref span, ref written);
        }

        "\"/>"u8.WriteTo(buffer, ref span, ref written);
    }

    private static void WriteColumns(
        IReadOnlyCollection<Column> columns,
        BuffersChain buffer,
        ref Span<byte> span,
        ref int written)
    {
        "<cols>"u8.WriteTo(buffer, ref span, ref written);

        var index = 1;

        foreach (var column in columns)
        {
            // column width will be applied to columns with indexes between min and max
            "<col min=\""u8.WriteTo(buffer, ref span, ref written);
            index.WriteTo(buffer, ref span, ref written);

            "\" max=\""u8.WriteTo(buffer, ref span, ref written);
            index.WriteTo(buffer, ref span, ref written);

            "\" width=\""u8.WriteTo(buffer, ref span, ref written);
            column.Width.WriteTo(buffer, ref span, ref written);

            "\" customWidth=\"1\"/>"u8.WriteTo(buffer, ref span, ref written);

            index++;
        }

        "</cols>"u8.WriteTo(buffer, ref span, ref written);
    }

    private static void WriteMerges(
        IReadOnlyCollection<Merge> merges,
        BuffersChain buffer,
        ref Span<byte> span,
        ref int written)
    {
        if (merges.Count == 0)
            return;

        "<mergeCells>"u8.WriteTo(buffer, ref span, ref written);

        foreach (var merge in merges)
        {
            "<mergeCell ref=\""u8.WriteTo(buffer, ref span, ref written);
            merge.WriteTo(buffer, ref span, ref written);
            "\"/>"u8.WriteTo(buffer, ref span, ref written);
        }

        "</mergeCells>"u8.WriteTo(buffer, ref span, ref written);
    }

    private static void WriteDrawing(
        Drawing drawing,
        BuffersChain buffer,
        Encoder encoder,
        ref Span<byte> span,
        ref int written)
    {
        if (drawing.IsEmpty)
            return;

        "<drawing r:id=\""u8.WriteTo(buffer, ref span, ref written);
        drawing.RelationshipId.WriteTo(buffer, encoder, ref span, ref written);
        "\"/>"u8.WriteTo(buffer, ref span, ref written);
    }

    private static void WriteHyperlinks(
        Dictionary<string, List<CellReference>> hyperlinks,
        BuffersChain buffer,
        ref Span<byte> span,
        ref int written)
    {
        if (hyperlinks.Count == 0)
            return;

        "<hyperlinks>"u8.WriteTo(buffer, ref span, ref written);

        var count = 0;

        foreach (var hyperlinkPair in hyperlinks)
        {
            foreach (var cellReference in hyperlinkPair.Value)
            {
                "<hyperlink r:id=\"link"u8.WriteTo(buffer, ref span, ref written);
                count.WriteTo(buffer, ref span, ref written);
                "\" ref=\""u8.WriteTo(buffer, ref span, ref written);
                cellReference.WriteTo(buffer, ref span, ref written);
                "\"/>"u8.WriteTo(buffer, ref span, ref written);
            }

            count++;
        }

        "</hyperlinks>"u8.WriteTo(buffer, ref span, ref written);
    }
}
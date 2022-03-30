using System;
using System.Collections.Generic;
using Gooseberry.ExcelStreaming.Configuration;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class SheetWriter
{
    public void WriteStartSheet(BuffersChain buffer, SheetConfiguration? configuration)
    {
        var span = buffer.GetSpan();
        var written = 0;
        
        Constants.Worksheet.Prefix.WriteTo(buffer, ref span, ref written);
            
        if (configuration?.TopLeftUnpinnedCell != null)
            WriteTopLeftUnpinnedCell(configuration.Value.TopLeftUnpinnedCell.Value, buffer, ref span, ref written);

        if (configuration?.Columns != null && configuration?.Columns.Count != 0)
            WriteColumns(configuration!.Value.Columns, buffer, ref span, ref written);
            
        Constants.Worksheet.SheetData.Prefix.WriteTo(buffer, ref span, ref written);
        
        buffer.Advance(written);
    }

    public void WriteEndSheet(
        BuffersChain buffer, 
        IReadOnlyCollection<Merge> merges, 
        Dictionary<string, List<CellReference>> hyperlinks)
    {
        var span = buffer.GetSpan();
        var written = 0;
     
        Constants.Worksheet.SheetData.Postfix.WriteTo(buffer, ref span, ref written);
            
        if (merges.Count != 0)
            WriteMerges(merges, buffer, ref span, ref written); 
            
        if (hyperlinks.Count != 0)
            WriteHyperlinks(hyperlinks, buffer, ref span, ref written);

        Constants.Worksheet.Postfix.WriteTo(buffer, ref span, ref written);        
        
        buffer.Advance(written);        
    }
    
    private static void WriteTopLeftUnpinnedCell(
        CellReference cellReference, 
        BuffersChain buffer, 
        ref Span<byte> span,
        ref int written)
    {
        Constants.Worksheet.View.Prefix.WriteTo(buffer, ref span, ref written);
            
        Constants.Worksheet.View.TopLeftCell.Prefix.WriteTo(buffer, ref span, ref written);
        cellReference.WriteTo(buffer, ref span, ref written);
        Constants.Worksheet.View.TopLeftCell.Postfix.WriteTo(buffer, ref span, ref written);
            
        Constants.Worksheet.View.YSplit.Prefix.WriteTo(buffer, ref span, ref written);
        (cellReference.Row - 1).WriteTo(buffer, ref span, ref written);
        Constants.Worksheet.View.YSplit.Postfix.WriteTo(buffer, ref span, ref written);
            
        Constants.Worksheet.View.XSplit.Prefix.WriteTo(buffer, ref span, ref written);
        (cellReference.Column - 1).WriteTo(buffer, ref span, ref written);
        Constants.Worksheet.View.XSplit.Postfix.WriteTo(buffer, ref span, ref written);

        Constants.Worksheet.View.Postfix.WriteTo(buffer, ref span, ref written);
    }

    private static void WriteColumns(
        IReadOnlyCollection<Column> columns,
        BuffersChain buffer, 
        ref Span<byte> span,
        ref int written)
    {
        Constants.Worksheet.Columns.Prefix.WriteTo(buffer, ref span, ref written);
            
        var index = 1;
        foreach (var column in columns)
        {
            // column width will be applied to columns with indexes between min and max
            Constants.Worksheet.Columns.Item.Prefix.WriteTo(buffer, ref span, ref written);
                
            Constants.Worksheet.Columns.Item.Min.Prefix.WriteTo(buffer, ref span, ref written);
            index.WriteTo(buffer, ref span, ref written);
            Constants.Worksheet.Columns.Item.Min.Postfix.WriteTo(buffer, ref span, ref written);
                
            Constants.Worksheet.Columns.Item.Max.Prefix.WriteTo(buffer, ref span, ref written);
            index.WriteTo(buffer, ref span, ref written);
            Constants.Worksheet.Columns.Item.Max.Postfix.WriteTo(buffer, ref span, ref written);

            Constants.Worksheet.Columns.Item.Width.Prefix.WriteTo(buffer, ref span, ref written);
            column.Width.WriteTo(buffer, ref span, ref written);
            Constants.Worksheet.Columns.Item.Width.Postfix.WriteTo(buffer, ref span, ref written);
                
            Constants.Worksheet.Columns.Item.Postfix.WriteTo(buffer, ref span, ref written);

            index++;
        }

        Constants.Worksheet.Columns.Postfix.WriteTo(buffer, ref span, ref written);        
    }

    private static void WriteMerges(
        IReadOnlyCollection<Merge> merges,
        BuffersChain buffer,
        ref Span<byte> span,
        ref int written)
    {
        Constants.Worksheet.Merges.Prefix.WriteTo(buffer, ref span, ref written);
            
        foreach (var merge in merges)
        {
            Constants.Worksheet.Merges.Merge.Prefix.WriteTo(buffer, ref span, ref written);
            merge.WriteTo(buffer, ref span, ref written);
            Constants.Worksheet.Merges.Merge.Postfix.WriteTo(buffer, ref span, ref written);
        }
            
        Constants.Worksheet.Merges.Postfix.WriteTo(buffer, ref span, ref written);
    }
    
    private static void WriteHyperlinks(
        Dictionary<string, List<CellReference>> hyperlinks,
        BuffersChain buffer,
        ref Span<byte> span,
        ref int written)
    {
        Constants.Worksheet.Hyperlinks.Prefix.WriteTo(buffer, ref span, ref written);

        var count = 0;
        foreach (var hyperlinkPair in hyperlinks)
        {
            foreach (var cellReference in hyperlinkPair.Value)
            {
                Constants.Worksheet.Hyperlinks.Hyperlink.StartPrefix.WriteTo(buffer, ref span, ref written);
                count.WriteTo(buffer, ref span, ref written);
                Constants.Worksheet.Hyperlinks.Hyperlink.EndPrefix.WriteTo(buffer, ref span, ref written);
                cellReference.WriteTo(buffer, ref span, ref written);
                Constants.Worksheet.Hyperlinks.Hyperlink.Postfix.WriteTo(buffer, ref span, ref written);
            }

            count++;
        }
            
        Constants.Worksheet.Hyperlinks.Postfix.WriteTo(buffer, ref span, ref written);
    }
    
}
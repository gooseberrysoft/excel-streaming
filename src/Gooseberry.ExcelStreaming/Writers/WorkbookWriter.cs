using System.Text;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct WorkbookWriter
{
    public void Write(IReadOnlyCollection<Sheet> sheets, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;
        
        Constants.Workbook.Prefix.WriteTo(buffer, ref span, ref written);

        foreach (var sheet in sheets)
        {
            Constants.Workbook.Sheet.StartPrefix.WriteTo(buffer, ref span, ref written);
            sheet.Name.WriteEscapedTo(buffer, encoder, ref span, ref written);
            Constants.Workbook.Sheet.EndPrefix.WriteTo(buffer, ref span, ref written);
                
            sheet.Id.WriteTo(buffer, ref span, ref written);
            Constants.Workbook.Sheet.EndPostfix.WriteTo(buffer, ref span, ref written);
            sheet.RelationshipId.WriteTo(buffer, encoder, ref span, ref written);
            Constants.Workbook.Sheet.Postfix.WriteTo(buffer, ref span, ref written);
        }

        Constants.Workbook.Postfix.WriteTo(buffer, ref span, ref written);
        
        buffer.Advance(written);        
    }
}
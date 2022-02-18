using System.Collections.Generic;
using System.Text;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct WorkbookRelationshipsWriter
{
    public void Write(IReadOnlyCollection<Sheet> sheets, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);
        Constants.WorkbookRelationships.Prefix.WriteTo(buffer, ref span, ref written);
                
        foreach (var sheet in sheets)
        {
            Constants.WorkbookRelationships.Sheet.Prefix.WriteTo(buffer, ref span, ref written);
            sheet.RelationshipId.WriteTo(buffer, encoder, ref span, ref written);
            Constants.WorkbookRelationships.Sheet.Middle.WriteTo(buffer, ref span, ref written);
            sheet.RelationshipId.WriteTo(buffer, encoder, ref span, ref written);
            Constants.WorkbookRelationships.Sheet.Postfix.WriteTo(buffer, ref span, ref written);
        }

        Constants.WorkbookRelationships.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
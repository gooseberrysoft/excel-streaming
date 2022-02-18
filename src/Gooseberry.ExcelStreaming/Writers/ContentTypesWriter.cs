using System.Collections.Generic;
using System.Text;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct ContentTypesWriter
{
    public void Write(IReadOnlyCollection<Sheet> sheets, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);
        Constants.ContentTypes.Prefix.WriteTo(buffer, ref span, ref written);
            
        foreach (var sheet in sheets)
        {
            Constants.ContentTypes.Sheet.Prefix.WriteTo(buffer, ref span, ref written);
            sheet.RelationshipId.WriteTo(buffer, encoder, ref span, ref written);
            Constants.ContentTypes.Sheet.Postfix.WriteTo(buffer, ref span, ref written);
        }

        Constants.ContentTypes.Postfix.WriteTo(buffer, ref span, ref written);
        
        buffer.Advance(written);        
    }
}
using System.Collections.Generic;
using System.Text;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct SheetRelationshipsWriter
{
    public void Write(IReadOnlyCollection<string> hyperlinks, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);
        Constants.SheetRelationships.Prefix.WriteTo(buffer, ref span, ref written);

        var count = 0;
        foreach (var hyperlink in hyperlinks)
        {
            Constants.SheetRelationships.HyperlinkRelationship.StartPrefix.WriteTo(buffer, ref span, ref written);
            count.WriteTo(buffer, ref span, ref written);
            Constants.SheetRelationships.HyperlinkRelationship.EndPrefix.WriteTo(buffer, ref span, ref written);
            hyperlink.WriteEscapedTo(buffer, encoder, ref span, ref written);
            Constants.SheetRelationships.HyperlinkRelationship.Postfix.WriteTo(buffer, ref span, ref written);
            
            count++;
        }

        Constants.SheetRelationships.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
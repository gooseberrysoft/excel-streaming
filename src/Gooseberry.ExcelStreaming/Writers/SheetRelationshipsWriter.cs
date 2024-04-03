using System.Collections.Generic;
using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct SheetRelationshipsWriter
{
    public void Write(Sheet sheet, IReadOnlyCollection<string> hyperlinks, Drawing drawing, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);
        Constants.SheetRelationships.Prefix.WriteTo(buffer, ref span, ref written);

        var count = 0;
        foreach (var hyperlink in hyperlinks)
        {
            Constants.SheetRelationships.Hyperlink.StartPrefix.WriteTo(buffer, ref span, ref written);
            count.WriteTo(buffer, ref span, ref written);
            Constants.SheetRelationships.Hyperlink.EndPrefix.WriteTo(buffer, ref span, ref written);
            hyperlink.WriteEscapedTo(buffer, encoder, ref span, ref written);
            Constants.SheetRelationships.Hyperlink.Postfix.WriteTo(buffer, ref span, ref written);
            
            count++;
        }

        if (!drawing.IsEmpty)
        {
            Constants.SheetRelationships.Drawing.GetPrefix().WriteTo(buffer, ref span, ref written);
            
            Constants.SheetRelationships.Drawing.Id.GetPrefix().WriteTo(buffer, ref span, ref written);
            drawing.RelationshipId.WriteTo(buffer, encoder, ref span, ref written);
            Constants.SheetRelationships.Drawing.Id.GetPostfix().WriteTo(buffer, ref span, ref written);
            
            Constants.SheetRelationships.Drawing.Target.GetPrefix().WriteTo(buffer, ref span, ref written);
            PathResolver.GetRelativePath(sheet, drawing).WriteTo(buffer, encoder, ref span, ref written);
            Constants.SheetRelationships.Drawing.Target.GetPostfix().WriteTo(buffer, ref span, ref written);
            
            Constants.SheetRelationships.Drawing.GetPostfix().WriteTo(buffer, ref span, ref written);
        }

        Constants.SheetRelationships.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}

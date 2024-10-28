using System.Text;
using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class DrawingRelationshipsWriter
{
    public static void Write(Drawing drawing, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);

        Constants.Drawing.Relationships.GetPrefix().WriteTo(buffer, ref span, ref written);

        foreach (var picture in drawing.Pictures)
        {
            Constants.Drawing.Relationships.Relationship.GetPrefix().WriteTo(buffer, ref span, ref written);

            Constants.Drawing.Relationships.Relationship.Target.GetPrefix().WriteTo(buffer, ref span, ref written);
            PathResolver.GetPictureFullPath(picture).EnsureLeadingSlash().WriteTo(buffer, encoder, ref span, ref written);
            Constants.Drawing.Relationships.Relationship.Target.GetPostfix().WriteTo(buffer, ref span, ref written);

            Constants.Drawing.Relationships.Relationship.Id.GetPrefix().WriteTo(buffer, ref span, ref written);
            picture.RelationshipId.WriteTo(buffer, encoder, ref span, ref written);
            Constants.Drawing.Relationships.Relationship.Id.GetPostfix().WriteTo(buffer, ref span, ref written);

            Constants.Drawing.Relationships.Relationship.GetPostfix().WriteTo(buffer, ref span, ref written);
        }

        Constants.Drawing.Relationships.GetPostfix().WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
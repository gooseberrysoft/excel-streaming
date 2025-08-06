using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class DrawingRelationshipsWriter
{
    private static ReadOnlySpan<byte> Prefix => "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"u8;

    private static ReadOnlySpan<byte> Postfix => "</Relationships>"u8;

    private static ReadOnlySpan<byte> RelationshipPrefix
        => "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\""u8;

    private static ReadOnlySpan<byte> RelationshipPostfix => "/>"u8;

    private static ReadOnlySpan<byte> IdPrefix => " Id=\""u8;

    private static ReadOnlySpan<byte> IdPostfix => "\""u8;

    private static ReadOnlySpan<byte> TargetPrefix => " Target=\"/"u8;

    private static ReadOnlySpan<byte> TargetPostfix => "\""u8;

    public static void Write(Drawing drawing, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);

        Prefix.WriteTo(buffer, ref span, ref written);

        foreach (var picture in drawing.Pictures)
        {
            RelationshipPrefix.WriteTo(buffer, ref span, ref written);

            TargetPrefix.WriteTo(buffer, ref span, ref written);
            PathResolver.GetPictureFullPath(picture).WriteTo(buffer, encoder, ref span, ref written);
            TargetPostfix.WriteTo(buffer, ref span, ref written);

            IdPrefix.WriteTo(buffer, ref span, ref written);
            picture.RelationshipId.WriteTo(buffer, encoder, ref span, ref written);
            IdPostfix.WriteTo(buffer, ref span, ref written);

            RelationshipPostfix.WriteTo(buffer, ref span, ref written);
        }

        Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
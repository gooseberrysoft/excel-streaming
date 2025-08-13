using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class SheetRelationshipsWriter
{
    private static ReadOnlySpan<byte> Prefix =>
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"u8;

    private static ReadOnlySpan<byte> Postfix => "</Relationships>"u8;

    private static ReadOnlySpan<byte> HyperlinkStartPrefix => "<Relationship Id=\"link"u8;

    private static ReadOnlySpan<byte> HyperlinkEndPrefix =>
        "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\""u8;

    private static ReadOnlySpan<byte> HyperlinkPostfix => "\" TargetMode=\"External\"/>"u8;

    private static ReadOnlySpan<byte> DrawingPrefix
        => "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\""u8;

    private static ReadOnlySpan<byte> DrawingPostfix => "/>"u8;
    private static ReadOnlySpan<byte> DrawingIdPrefix => " Id=\""u8;
    private static ReadOnlySpan<byte> DrawingIdPostfix => "\""u8;
    private static ReadOnlySpan<byte> DrawingTargetPrefix => " Target=\"/"u8;
    private static ReadOnlySpan<byte> DrawingTargetPostfix => "\""u8;

    public static void Write(IReadOnlyCollection<string>? hyperlinks, Drawing drawing, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);
        Prefix.WriteTo(buffer, ref span, ref written);

        if (hyperlinks != null)
        {
            var count = 0;
            foreach (var hyperlink in hyperlinks)
            {
                HyperlinkStartPrefix.WriteTo(buffer, ref span, ref written);
                count.WriteTo(buffer, ref span, ref written);
                HyperlinkEndPrefix.WriteTo(buffer, ref span, ref written);
                hyperlink.WriteEscapedTo(buffer, encoder, ref span, ref written);
                HyperlinkPostfix.WriteTo(buffer, ref span, ref written);

                count++;
            }
        }

        if (!drawing.IsEmpty)
        {
            DrawingPrefix.WriteTo(buffer, ref span, ref written);

            DrawingTargetPrefix.WriteTo(buffer, ref span, ref written);
            PathResolver.GetDrawingFullPath(drawing).WriteTo(buffer, encoder, ref span, ref written);
            DrawingTargetPostfix.WriteTo(buffer, ref span, ref written);

            DrawingIdPrefix.WriteTo(buffer, ref span, ref written);
            drawing.RelationshipId.WriteTo(buffer, encoder, ref span, ref written);
            DrawingIdPostfix.WriteTo(buffer, ref span, ref written);

            DrawingPostfix.WriteTo(buffer, ref span, ref written);
        }

        Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
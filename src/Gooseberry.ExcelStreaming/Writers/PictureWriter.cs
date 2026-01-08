using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class PictureWriter
{
    public static void Write(Picture picture, BufferSequence buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Write(picture, buffer, encoder, ref span, ref written);

        buffer.Advance(written);
    }

    public static void Write(Picture picture, BufferSequence buffer, Encoder encoder, ref Span<byte> span, ref int written)
    {
        "<xdr:pic>"u8.WriteTo(buffer, ref span, ref written);

        WriteNonVisualProperties(picture, buffer, encoder, ref span, ref written);
        WriteBinaryLargeImage(picture, buffer, encoder, ref span, ref written);
        WriteShapeProperties(buffer, ref span, ref written);

        "</xdr:pic>"u8.WriteTo(buffer, ref span, ref written);
    }

    private static void WriteShapeProperties(BufferSequence buffer, ref Span<byte> span, ref int written)
    {
        "<xdr:spPr><a:prstGeom prst=\"rect\"/></xdr:spPr>"u8.WriteTo(buffer, ref span, ref written);
    }

    private static void WriteBinaryLargeImage(Picture picture, BufferSequence buffer, Encoder encoder, ref Span<byte> span, ref int written)
    {
        "<xdr:blipFill><a:blip r:embed=\""u8.WriteTo(buffer, ref span, ref written);

        picture.RelationshipId.WriteTo(buffer, encoder, ref span, ref written);

        "\" cstate=\"print\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"u8.WriteTo(buffer, ref span, ref written);
    }

    private static void WriteNonVisualProperties(
        Picture picture,
        BufferSequence buffer,
        Encoder encoder,
        ref Span<byte> span,
        ref int written)
    {
        "<xdr:nvPicPr><xdr:cNvPr  id=\""u8.WriteTo(buffer, ref span, ref written);

        Utf8SpanFormattableWriter.WriteValue(picture.Id, buffer, ref span, ref written);

        "\" name=\""u8.WriteTo(buffer, ref span, ref written);

        picture.Name.WriteTo(buffer, encoder, ref span, ref written);

        "\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>"u8.WriteTo(buffer, ref span, ref written);
    }
}
using System.Drawing;
using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class OneCellAnchorPicturePlacementWriter(AnchorCell from, Size size) : IPicturePlacementWriter
{
    public void Write(Picture picture, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Write(picture, buffer, encoder, ref span, ref written);

        buffer.Advance(written);
    }

    public void Write(Picture picture, BuffersChain buffer, Encoder encoder, ref Span<byte> span, ref int written)
    {
        "<xdr:oneCellAnchor><xdr:from>"u8.WriteTo(buffer, ref span, ref written);

        AnchorCellWriter.Write(from, buffer, ref span, ref written);
        "</xdr:from><xdr:ext cy=\""u8.WriteTo(buffer, ref span, ref written);

        Utf8SpanFormattableWriter.WriteValue(EmuConverter.ConvertToEnglishMetricUnits(size.Width, resolution: 96),
            buffer, ref span, ref written);

        "\" cx=\""u8.WriteTo(buffer, ref span, ref written);

        Utf8SpanFormattableWriter.WriteValue(EmuConverter.ConvertToEnglishMetricUnits(size.Height, resolution: 96),
            buffer, ref span, ref written);
        "\"/>"u8.WriteTo(buffer, ref span, ref written);

        PictureWriter.Write(picture, buffer, encoder, ref span, ref written);

        "<xdr:clientData/></xdr:oneCellAnchor>"u8.WriteTo(buffer, ref span, ref written);
    }
}
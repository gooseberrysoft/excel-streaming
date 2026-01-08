using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class TwoCellAnchorPicturePlacementWriter(AnchorCell from, AnchorCell to) : IPicturePlacementWriter
{
    public void Write(Picture picture, BufferSequence buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Write(picture, buffer, encoder, ref span, ref written);

        buffer.Advance(written);
    }

    public void Write(Picture picture, BufferSequence buffer, Encoder encoder, ref Span<byte> span, ref int written)
    {
        "<xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from>"u8.WriteTo(buffer, ref span, ref written);

        AnchorCellWriter.Write(from, buffer, ref span, ref written);

        "</xdr:from><xdr:to>"u8.WriteTo(buffer, ref span, ref written);

        AnchorCellWriter.Write(to, buffer, ref span, ref written);

        "</xdr:to>"u8.WriteTo(buffer, ref span, ref written);

        PictureWriter.Write(picture, buffer, encoder, ref span, ref written);

        "<xdr:clientData/></xdr:twoCellAnchor>"u8.WriteTo(buffer, ref span, ref written);
    }
}
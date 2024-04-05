using System.Text;
using Gooseberry.ExcelStreaming.Pictures;
using Gooseberry.ExcelStreaming.Pictures.Placements;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class TwoCellAnchorPicturePlacementWriter
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
        if (picture.Placement is not TwoCellAnchorPicturePlacement placement)
        {
            throw new ArgumentException("Must be one cell anchor.", nameof(picture));
        }

        Constants.Drawing.TwoCellAnchor.GetPrefix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.AnchorFrom.GetPrefix().WriteTo(buffer, ref span, ref written);
        DataWriters.AnchorCellWriter.Write(placement.From, buffer, ref span, ref written);
        Constants.Drawing.AnchorFrom.GetPostfix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.AnchorTo.GetPrefix().WriteTo(buffer, ref span, ref written);
        DataWriters.AnchorCellWriter.Write(placement.To, buffer, ref span, ref written);
        Constants.Drawing.AnchorTo.GetPostfix().WriteTo(buffer, ref span, ref written);

        DataWriters.PictureWriter.Write(picture, buffer, encoder, ref span, ref written);

        Constants.Drawing.ClientData.GetBody().WriteTo(buffer, ref span, ref written);
        
        Constants.Drawing.TwoCellAnchor.GetPostfix().WriteTo(buffer, ref span, ref written);
    }
}

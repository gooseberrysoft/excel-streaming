using System.Text;
using Gooseberry.ExcelStreaming.Helpers;
using Gooseberry.ExcelStreaming.Pictures;
using Gooseberry.ExcelStreaming.Pictures.Placements;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class OneCellAnchorPicturePlacementWriter
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
        if (picture.Placement is not OneCellAnchorPicturePlacement placement)
        {
            throw new ArgumentException("Must be one cell anchor.", nameof(picture));
        }

        Constants.Drawing.OneCellAnchor.GetPrefix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.AnchorFrom.GetPrefix().WriteTo(buffer, ref span, ref written);
        DataWriters.AnchorCellWriter.Write(placement.From, buffer, ref span, ref written);
        Constants.Drawing.AnchorFrom.GetPostfix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.OneCellAnchor.Size.GetPrefix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.OneCellAnchor.Size.Width.GetPrefix().WriteTo(buffer, ref span, ref written);
        EmuConverter.ConvertToEnglishMetricUnits(placement.Size.Width, resolution: 96).WriteTo(buffer, ref span, ref written);
        Constants.Drawing.OneCellAnchor.Size.Width.GetPostfix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.OneCellAnchor.Size.Height.GetPrefix().WriteTo(buffer, ref span, ref written);
        EmuConverter.ConvertToEnglishMetricUnits(placement.Size.Height, resolution: 96).WriteTo(buffer, ref span, ref written);
        Constants.Drawing.OneCellAnchor.Size.Height.GetPostfix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.OneCellAnchor.Size.GetPostfix().WriteTo(buffer, ref span, ref written);

        DataWriters.PictureWriter.Write(picture, buffer, encoder, ref span, ref written);

        Constants.Drawing.ClientData.GetBody().WriteTo(buffer, ref span, ref written);
        
        Constants.Drawing.OneCellAnchor.GetPostfix().WriteTo(buffer, ref span, ref written);
    }
}

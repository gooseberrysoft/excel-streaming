using System.Text;
using Gooseberry.ExcelStreaming.Pictures;
using Gooseberry.ExcelStreaming.Pictures.Placements;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class DrawingWriter
{
    public void Write(Drawing drawing, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);
        Constants.Drawing.GetPrefix().WriteTo(buffer, ref span, ref written);

        foreach (var picture in drawing.Pictures)
        {
            switch (picture.Placement)
            {
                case OneCellAnchorPicturePlacement:
                    DataWriters.OneCellAnchorPicturePlacementWriter.Write(picture, buffer, encoder, ref span, ref written);

                    break;
                
                case TwoCellAnchorPicturePlacement:
                    DataWriters.TwoCellAnchorPicturePlacementWriter.Write(picture, buffer, encoder, ref span, ref written);

                    break;
            }
        }
        
        Constants.Drawing.GetPostfix().WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}

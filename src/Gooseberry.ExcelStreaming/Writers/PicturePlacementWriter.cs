using System.Text;
using Gooseberry.ExcelStreaming.Helpers;
using Gooseberry.ExcelStreaming.Pictures;
using Gooseberry.ExcelStreaming.Pictures.Abstractions;
using Gooseberry.ExcelStreaming.Pictures.Placements;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct PicturePlacementWriter : IPicturePlacementVisitor
{
    private readonly Picture _picture;
    private readonly Encoder _encoder;
    private readonly BuffersChain _buffer;

    public PicturePlacementWriter(
        BuffersChain buffer,
        Encoder encoder,
        Picture picture)
    {
        _buffer = buffer;
        _encoder = encoder;
        _picture = picture;
    }

    public void Write()
        => _picture.Placement.Visit(this);

    void IPicturePlacementVisitor.Visit(OneCellAnchorPicturePlacement placement)
    {
        var span = _buffer.GetSpan();
        var written = 0;
        
        Constants.Drawing.OneCellAnchor.GetPrefix().WriteTo(_buffer, ref span, ref written);

        Constants.Drawing.AnchorFrom.GetPrefix().WriteTo(_buffer, ref span, ref written);
        DataWriters.AnchorCellWriter.Write(placement.From, _buffer, ref span, ref written);
        Constants.Drawing.AnchorFrom.GetPostfix().WriteTo(_buffer, ref span, ref written);

        Constants.Drawing.OneCellAnchor.Size.GetPrefix().WriteTo(_buffer, ref span, ref written);

        Constants.Drawing.OneCellAnchor.Size.Width.GetPrefix().WriteTo(_buffer, ref span, ref written);
        EmuConverter.ConvertToEnglishMetricUnits(placement.Size.Width, resolution: 96).WriteTo(_buffer, ref span, ref written);
        Constants.Drawing.OneCellAnchor.Size.Width.GetPostfix().WriteTo(_buffer, ref span, ref written);

        Constants.Drawing.OneCellAnchor.Size.Height.GetPrefix().WriteTo(_buffer, ref span, ref written);
        EmuConverter.ConvertToEnglishMetricUnits(placement.Size.Height, resolution: 96).WriteTo(_buffer, ref span, ref written);
        Constants.Drawing.OneCellAnchor.Size.Height.GetPostfix().WriteTo(_buffer, ref span, ref written);

        Constants.Drawing.OneCellAnchor.Size.GetPostfix().WriteTo(_buffer, ref span, ref written);

        DataWriters.PictureWriter.Write(_picture, _buffer, _encoder, ref span, ref written);

        Constants.Drawing.ClientData.GetBody().WriteTo(_buffer, ref span, ref written);

        Constants.Drawing.OneCellAnchor.GetPostfix().WriteTo(_buffer, ref span, ref written);

        _buffer.Advance(written);
    }

    void IPicturePlacementVisitor.Visit(TwoCellAnchorPicturePlacement placement)
    {
        var span = _buffer.GetSpan();
        var written = 0;

        Constants.Drawing.TwoCellAnchor.GetPrefix().WriteTo(_buffer, ref span, ref written);

        Constants.Drawing.AnchorFrom.GetPrefix().WriteTo(_buffer, ref span, ref written);
        DataWriters.AnchorCellWriter.Write(placement.From, _buffer, ref span, ref written);
        Constants.Drawing.AnchorFrom.GetPostfix().WriteTo(_buffer, ref span, ref written);

        Constants.Drawing.AnchorTo.GetPrefix().WriteTo(_buffer, ref span, ref written);
        DataWriters.AnchorCellWriter.Write(placement.To, _buffer, ref span, ref written);
        Constants.Drawing.AnchorTo.GetPostfix().WriteTo(_buffer, ref span, ref written);

        DataWriters.PictureWriter.Write(_picture, _buffer, _encoder, ref span, ref written);

        Constants.Drawing.ClientData.GetBody().WriteTo(_buffer, ref span, ref written);

        Constants.Drawing.TwoCellAnchor.GetPostfix().WriteTo(_buffer, ref span, ref written);

        _buffer.Advance(written);
    }
}

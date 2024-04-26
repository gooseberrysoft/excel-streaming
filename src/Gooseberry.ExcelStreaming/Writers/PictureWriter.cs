using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class PictureWriter
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
        Constants.Drawing.Picture.GetPrefix().WriteTo(buffer, ref span, ref written);

        WriteNonVisualProperties(picture, buffer, encoder, ref span, ref written);
        WriteBinaryLargeImage(picture, buffer, encoder, ref span, ref written);
        WriteShapeProperties(buffer, ref span, ref written);

        Constants.Drawing.Picture.GetPostfix().WriteTo(buffer, ref span, ref written);
    }

    private static void WriteShapeProperties(BuffersChain buffer, ref Span<byte> span, ref int written)
    {
        Constants.Drawing.Picture.ShapeProperties.GetPrefix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.Picture.ShapeProperties.PresetGeometry.GetRect().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.Picture.ShapeProperties.GetPostfix().WriteTo(buffer, ref span, ref written);
    }

    private static void WriteBinaryLargeImage(Picture picture, BuffersChain buffer, Encoder encoder, ref Span<byte> span, ref int written)
    {
        Constants.Drawing.Picture.BlipFill.GetPrefix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.Picture.BlipFill.Blip.GetPrefix().WriteTo(buffer, ref span, ref written);
        picture.RelationshipId.WriteTo(buffer, encoder, ref span, ref written);
        Constants.Drawing.Picture.BlipFill.Blip.GetPostfix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.Picture.BlipFill.Stretch.GetFillRect().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.Picture.BlipFill.GetPostfix().WriteTo(buffer, ref span, ref written);
    }

    private static void WriteNonVisualProperties(
        Picture picture,
        BuffersChain buffer,
        Encoder encoder,
        ref Span<byte> span,
        ref int written)
    {
        Constants.Drawing.Picture.NonVisualProperties.GetPrefix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.Picture.NonVisualProperties.Properties.GetPrefix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.Picture.NonVisualProperties.Properties.Id.GetPrefix().WriteTo(buffer, ref span, ref written);
        picture.Id.WriteTo(buffer, ref span, ref written);
        Constants.Drawing.Picture.NonVisualProperties.Properties.Id.GetPostfix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.Picture.NonVisualProperties.Properties.Name.GetPrefix().WriteTo(buffer, ref span, ref written);
        picture.Name.WriteTo(buffer, encoder, ref span, ref written);
        Constants.Drawing.Picture.NonVisualProperties.Properties.Name.GetPostfix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.Picture.NonVisualProperties.Properties.GetPostfix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.Picture.NonVisualProperties.PictureProperties.GetBody().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.Picture.NonVisualProperties.GetPostfix().WriteTo(buffer, ref span, ref written);
    }
}
using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class DrawingWriter
{
    public static void Write(Drawing drawing, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);
        Constants.Drawing.GetPrefix().WriteTo(buffer, ref span, ref written);

        foreach (var picture in drawing.Pictures)
            picture.PlacementWriter.Write(picture, buffer, encoder, ref span, ref written);

        Constants.Drawing.GetPostfix().WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
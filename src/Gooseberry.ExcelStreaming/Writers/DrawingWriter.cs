using System.Text;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class DrawingWriter
{
    public void Write(Drawing drawing, BuffersChain buffer, Encoder encoder)
    {
        WritePrefix(buffer);

        foreach (var picture in drawing.Pictures)
            new PicturePlacementWriter(buffer, encoder, picture).Write();

        WritePostfix(buffer);
    }

    private static void WritePrefix(BuffersChain buffer)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);
        Constants.Drawing.GetPrefix().WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }

    private static void WritePostfix(BuffersChain buffer)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.Drawing.GetPostfix().WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}

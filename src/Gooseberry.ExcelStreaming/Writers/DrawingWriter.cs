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
        "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"u8
            .WriteTo(buffer, ref span, ref written);

        foreach (var picture in drawing.Pictures)
            picture.PlacementWriter.Write(picture, buffer, encoder, ref span, ref written);

        "</xdr:wsDr>"u8.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
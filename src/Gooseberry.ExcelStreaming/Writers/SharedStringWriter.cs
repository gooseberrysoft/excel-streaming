using System.Text;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class SharedStringWriter
{
    public static readonly byte[] EmptyTable =
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></sst>"u8.ToArray();

    public static void WritePrefix(BufferSequence buffer)
        => "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8.WriteTo(buffer);

    public static void Write(string value, BufferSequence buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        "<si><t>"u8.WriteTo(buffer, ref span, ref written);
        value.WriteEscapedTo(buffer, encoder, ref span, ref written);
        "</t></si>"u8.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }

    public static void WritePostfix(BufferSequence buffer)
        => "</sst>"u8.WriteTo(buffer);
}
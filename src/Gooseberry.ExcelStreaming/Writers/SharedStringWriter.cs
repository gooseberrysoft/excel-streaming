using System.Text;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class SharedStringWriter
{
    public static void WritePrefix(BuffersChain buffer)
        => Constants.SharedStringTable.Prefix.WriteTo(buffer);

    public static void Write(string value, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.SharedStringTable.Item.Prefix.WriteTo(buffer, ref span, ref written);
        value.WriteEscapedTo(buffer, encoder, ref span, ref written);
        Constants.SharedStringTable.Item.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }

    public static void WritePostfix(BuffersChain buffer)
        => Constants.SharedStringTable.Postfix.WriteTo(buffer);
}
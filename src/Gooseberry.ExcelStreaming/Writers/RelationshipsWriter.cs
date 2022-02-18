namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct RelationshipsWriter
{
    public void Write(BuffersChain buffer)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);
        Constants.Relationships.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
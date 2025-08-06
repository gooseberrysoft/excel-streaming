using System.Text;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class WorkbookRelationshipsWriter
{
    private static ReadOnlySpan<byte> WorkbookPrefix
        => "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"u8;

    private static ReadOnlySpan<byte> WorkbookPostfix =>
        """
        <Relationship Id="styles1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
        <Relationship Id="sharedStrings1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />
        </Relationships>
        """u8;
    
    private static ReadOnlySpan<byte> Prefix
        => "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"/xl/worksheets/sheet"u8;

    private static ReadOnlySpan<byte> Middle => ".xml\" Id=\"sheet"u8;
    private static ReadOnlySpan<byte> Postfix => "\"/>"u8;

    public static void Write(IReadOnlyCollection<Sheet> sheets, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);
        WorkbookPrefix.WriteTo(buffer, ref span, ref written);

        foreach (var sheet in sheets)
        {
            Prefix.WriteTo(buffer, ref span, ref written);
            sheet.Id.WriteTo(buffer, ref span, ref written);
            Middle.WriteTo(buffer, ref span, ref written);
            sheet.Id.WriteTo(buffer, ref span, ref written);
            Postfix.WriteTo(buffer, ref span, ref written);
        }

        WorkbookPostfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
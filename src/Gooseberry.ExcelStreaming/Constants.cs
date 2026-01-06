// ReSharper disable once CheckNamespace

namespace Gooseberry.ExcelStreaming;

internal static class Constants
{
    public static ReadOnlySpan<byte> XmlPrefix => "<?xml version=\"1.0\" encoding=\"utf-8\"?>"u8;

    public static byte[] RelationshipsContent =
        """
            <?xml version="1.0" encoding="utf-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/xl/workbook.xml" Id="R2196c6c3552b4024" />
            </Relationships>
            """u8.ToArray();
}
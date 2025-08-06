// ReSharper disable once CheckNamespace

namespace Gooseberry.ExcelStreaming;

internal static partial class Constants
{
    public static ReadOnlySpan<byte> TrueValue => "true"u8;

    public static ReadOnlySpan<byte> FalseValue => "false"u8;

    public static ReadOnlySpan<byte> XmlPrefix => "<?xml version=\"1.0\" encoding=\"utf-8\"?>"u8;

    public static class Worksheet
    {
        public static class SheetData
        {
            public static class Row
            {
                public static class Open
                {
                    public static ReadOnlySpan<byte> Prefix => "<row"u8;

                    public static ReadOnlySpan<byte> Postfix => ">"u8;

                    public static class Height
                    {
                        public static ReadOnlySpan<byte> Prefix => " ht=\""u8;

                        public static ReadOnlySpan<byte> Postfix => "\" customHeight=\"1\""u8;
                    }

                    public static class OutlineLevel
                    {
                        public static ReadOnlySpan<byte> Prefix => " outlineLevel=\""u8;

                        public static ReadOnlySpan<byte> Postfix => "\""u8;
                    }

                    public static ReadOnlySpan<byte> Collapsed => " collapsed=\"1\""u8;

                    public static ReadOnlySpan<byte> Hidden => " hidden=\"true\""u8;
                }

                public static ReadOnlySpan<byte> Postfix => "</row>"u8;

                public static class Cell
                {
                    public static ReadOnlySpan<byte> StringDataType => " t=\"str\""u8;

                    public static ReadOnlySpan<byte> SharedStringDataType => " t=\"s\""u8;

                    public static ReadOnlySpan<byte> NumberDataType => " t=\"n\""u8;

                    public static ReadOnlySpan<byte> DateTimeDataType => ""u8;

                    public static ReadOnlySpan<byte> Prefix => "<c"u8;

                    public static ReadOnlySpan<byte> Middle => "><v>"u8;

                    public static ReadOnlySpan<byte> Postfix => "</v></c>"u8;

                    public static class Style
                    {
                        public static ReadOnlySpan<byte> Prefix => " s=\""u8;

                        public static ReadOnlySpan<byte> Postfix => "\""u8;
                    }

                    public static ReadOnlySpan<byte> Empty => "<c t=\"str\"><v></v></c>"u8;
                }
            }
        }
    }

    public static class SheetRelationships
    {
        public static ReadOnlySpan<byte> Prefix =>
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"u8;

        public static ReadOnlySpan<byte> Postfix => "</Relationships>"u8;

        public static class Hyperlink
        {
            public static ReadOnlySpan<byte> StartPrefix => "<Relationship Id=\"link"u8;

            public static ReadOnlySpan<byte> EndPrefix =>
                "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\""u8;

            public static ReadOnlySpan<byte> Postfix => "\" TargetMode=\"External\"/>"u8;
        }

        public static class Drawing
        {
            public static ReadOnlySpan<byte> GetPrefix()
                => "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\""u8;

            public static ReadOnlySpan<byte> GetPostfix()
                => "/>"u8;

            public static class Id
            {
                public static ReadOnlySpan<byte> GetPrefix()
                    => " Id=\""u8;

                public static ReadOnlySpan<byte> GetPostfix()
                    => "\""u8;
            }

            public static class Target
            {
                public static ReadOnlySpan<byte> GetPrefix()
                    => " Target=\"/"u8;

                public static ReadOnlySpan<byte> GetPostfix()
                    => "\""u8;
            }
        }
    }

    public static byte[] RelationshipsContent =
        """
            <?xml version="1.0" encoding="utf-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/xl/workbook.xml" Id="R2196c6c3552b4024" />
            </Relationships>
            """u8.ToArray();

    public static class SharedStringTable
    {
        public static ReadOnlySpan<byte> Prefix => "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8;

        public static ReadOnlySpan<byte> Postfix => "</sst>"u8;

        public static readonly byte[] EmptyTable =
            "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></sst>"u8.ToArray();

        public static class Item
        {
            public static ReadOnlySpan<byte> Prefix => "<si><t>"u8;

            public static ReadOnlySpan<byte> Postfix => "</t></si>"u8;
        }
    }
}
// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal static partial class Constants
{
    public const double DefaultBufferFlushThreshold = 0.9;

    public static ReadOnlySpan<byte> TrueValue => "true"u8;

    public static ReadOnlySpan<byte> FalseValue => "false"u8;

    public static ReadOnlySpan<byte> XmlPrefix => "<?xml version=\"1.0\" encoding=\"utf-8\"?>"u8;

    public static class Workbook
    {
        public static ReadOnlySpan<byte> Prefix
            => "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheets>"u8;

        public static ReadOnlySpan<byte> Postfix => "</sheets></workbook>"u8;

        public static class Sheet
        {
            public static ReadOnlySpan<byte> StartPrefix => "<sheet name=\""u8;

            public static ReadOnlySpan<byte> EndPrefix => "\" sheetId=\""u8;

            public static ReadOnlySpan<byte> EndPostfix => "\" r:id=\""u8;

            public static ReadOnlySpan<byte> Postfix =>
                "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>"u8;
        }
    }

    public static class Worksheet
    {
        public static ReadOnlySpan<byte> Prefix =>
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"  xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"u8;

        public static ReadOnlySpan<byte> Postfix => "</worksheet>"u8;

        public static class Columns
        {
            public static ReadOnlySpan<byte> Prefix => "<cols>"u8;

            public static ReadOnlySpan<byte> Postfix => "</cols>"u8;

            public static class Item
            {
                public static ReadOnlySpan<byte> Prefix => "<col"u8;

                public static ReadOnlySpan<byte> Postfix => " customWidth=\"1\"/>"u8;

                public static class Min
                {
                    public static ReadOnlySpan<byte> Prefix => " min=\""u8;

                    public static ReadOnlySpan<byte> Postfix => "\""u8;
                }

                public static class Max
                {
                    public static ReadOnlySpan<byte> Prefix => " max=\""u8;

                    public static ReadOnlySpan<byte> Postfix => "\""u8;
                }

                public static class Width
                {
                    public static ReadOnlySpan<byte> Prefix => " width=\""u8;

                    public static ReadOnlySpan<byte> Postfix => "\""u8;
                }
            }
        }

        public static class Merges
        {
            public static ReadOnlySpan<byte> Prefix => "<mergeCells>"u8;

            public static ReadOnlySpan<byte> Postfix => "</mergeCells>"u8;

            public static class Merge
            {
                public static ReadOnlySpan<byte> Prefix => "<mergeCell ref=\""u8;

                public static ReadOnlySpan<byte> Postfix => "\"/>"u8;
            }
        }

        public static class Hyperlinks
        {
            public static ReadOnlySpan<byte> Prefix => "<hyperlinks>"u8;

            public static ReadOnlySpan<byte> Postfix => "</hyperlinks>"u8;

            public static class Hyperlink
            {
                public static ReadOnlySpan<byte> StartPrefix => "<hyperlink r:id=\"link"u8;

                public static ReadOnlySpan<byte> EndPrefix => "\" ref=\""u8;

                public static ReadOnlySpan<byte> Postfix => "\"/>"u8;
            }
        }

        public static class Drawings
        {
            public static ReadOnlySpan<byte> GetPrefix()
                => "<drawing r:id=\""u8;

            public static ReadOnlySpan<byte> GetPostfix()
                => "\"/>"u8;
        }

        public static class View
        {
            public static ReadOnlySpan<byte> Prefix => "<sheetViews><sheetView workbookViewId=\"0\""u8;

            public static ReadOnlySpan<byte> Middle => ">"u8;

            public static ReadOnlySpan<byte> Postfix => "</sheetView></sheetViews>"u8;

            public static class ShowGridLines
            {
                public static ReadOnlySpan<byte> Prefix => " showGridLines=\""u8;

                public static ReadOnlySpan<byte> Postfix => "\""u8;
            }

            public static class Pane
            {
                public static ReadOnlySpan<byte> Prefix => "<pane"u8;

                public static ReadOnlySpan<byte> Postfix => " activePane=\"bottomRight\" state=\"frozen\"/>"u8;

                public static class TopLeftCell
                {
                    public static ReadOnlySpan<byte> Prefix => " topLeftCell=\""u8;

                    public static ReadOnlySpan<byte> Postfix => "\""u8;
                }

                public static class YSplit
                {
                    public static ReadOnlySpan<byte> Prefix => " ySplit=\""u8;

                    public static ReadOnlySpan<byte> Postfix => "\""u8;
                }

                public static class XSplit
                {
                    public static ReadOnlySpan<byte> Prefix => " xSplit=\""u8;

                    public static ReadOnlySpan<byte> Postfix => "\""u8;
                }
            }
        }

        public static class SheetData
        {
            public static ReadOnlySpan<byte> Prefix => "<sheetData>"u8;

            public static ReadOnlySpan<byte> Postfix => "</sheetData>"u8;

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

    public static class ContentTypes
    {
        public static ReadOnlySpan<byte> Prefix =>
            """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
             <Default Extension ="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
             <Override PartName ="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
             <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml" />
            """u8;


        public static ReadOnlySpan<byte> Get(PictureFormat format)
        {
            return format switch
            {
                PictureFormat.Bmp => GetBmp(),
                PictureFormat.Gif => GetGif(),
                PictureFormat.Png => GetPng(),
                PictureFormat.Tiff => GetTiff(),
                PictureFormat.Icon => GetIcon(),
                PictureFormat.Jpeg => GetJpeg(),
                PictureFormat.Emf => GetEmf(),
                PictureFormat.Wmf => GetWmf(),
                _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unknown picture format")
            };
        }

        public static ReadOnlySpan<byte> GetBmp()
            => "<Default Extension=\"bmp\" ContentType=\"image/bmp\"/>"u8;

        public static ReadOnlySpan<byte> GetGif()
            => "<Default Extension=\"gif\" ContentType=\"image/gif\"/>"u8;

        public static ReadOnlySpan<byte> GetPng()
            => "<Default Extension=\"png\" ContentType=\"image/png\"/>"u8;

        public static ReadOnlySpan<byte> GetTiff()
            => "<Default Extension=\"tiff\" ContentType=\"image/tiff\"/>"u8;

        public static ReadOnlySpan<byte> GetIcon()
            => "<Default Extension=\"icon\" ContentType=\"image/x-icon\"/>"u8;

        public static ReadOnlySpan<byte> GetJpeg()
            => "<Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>"u8;

        public static ReadOnlySpan<byte> GetEmf()
            => "<Default Extension=\"emf\" ContentType=\"image/x-emf\"/>"u8;

        public static ReadOnlySpan<byte> GetWmf()
            => "<Default Extension=\"wmf\" ContentType=\"image/x-wmf\"/>"u8;

        public static ReadOnlySpan<byte> Postfix => "</Types>"u8;

        public static class Drawing
        {
            public static ReadOnlySpan<byte> GetPrefix()
                => "<Override PartName=\""u8;

            public static ReadOnlySpan<byte> GetPostfix()
                => "\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>"u8;
        }

        public static class Sheet
        {
            public static ReadOnlySpan<byte> Prefix => "<Override PartName=\"/xl/worksheets/"u8;

            public static ReadOnlySpan<byte> Postfix =>
                ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"u8;
        }
    }

    public static class WorkbookRelationships
    {
        public static ReadOnlySpan<byte> Prefix
            => "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"u8;

        public static ReadOnlySpan<byte> Postfix =>
            """
            <Relationship Id="styles1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
            <Relationship Id="sharedStrings1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />
            </Relationships>
            """u8;

        public static class Sheet
        {
            public static ReadOnlySpan<byte> Prefix
                => "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"/xl/worksheets/"u8;

            public static ReadOnlySpan<byte> Middle => ".xml\" Id=\""u8;
            public static ReadOnlySpan<byte> Postfix => "\"/>"u8;
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
                    => " Target=\""u8;

                public static ReadOnlySpan<byte> GetPostfix()
                    => "\""u8;
            }
        }
    }

    public static ReadOnlySpan<byte> Relationships =>
        """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/xl/workbook.xml" Id="R2196c6c3552b4024" />
        </Relationships>
        """u8;

    public static class SharedStringTable
    {
        public static ReadOnlySpan<byte> Prefix => "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8;

        public static ReadOnlySpan<byte> Postfix => "</sst>"u8;

        public static readonly byte[] EmptyTable = "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></sst>"u8.ToArray();

        public static class Item
        {
            public static ReadOnlySpan<byte> Prefix => "<si><t>"u8;

            public static ReadOnlySpan<byte> Postfix => "</t></si>"u8;
        }
    }
}
// ReSharper disable MemberHidesStaticFromOuterClass

namespace Gooseberry.ExcelStreaming;

internal static partial class Constants
{
    public static class Drawing
    {
        public static ReadOnlySpan<byte> GetPrefix()
            => "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"u8;

        public static ReadOnlySpan<byte> GetPostfix()
            => "</xdr:wsDr>"u8;

        public static class ClientData
        {
            public static ReadOnlySpan<byte> GetBody()
                => "<xdr:clientData/>"u8;
        }

        public static class Relationships
        {
            public static ReadOnlySpan<byte> GetPrefix()
                => "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"u8;

            public static ReadOnlySpan<byte> GetPostfix()
                => "</Relationships>"u8;

            public static class Relationship
            {
                public static ReadOnlySpan<byte> GetPrefix()
                    => "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\""u8;

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

        public static class Picture
        {
            public static ReadOnlySpan<byte> GetPrefix()
                => "<xdr:pic>"u8;

            public static ReadOnlySpan<byte> GetPostfix()
                => "</xdr:pic>"u8;

            public static class BlipFill
            {
                public static ReadOnlySpan<byte> GetPrefix()
                    => "<xdr:blipFill>"u8;

                public static ReadOnlySpan<byte> GetPostfix()
                    => "</xdr:blipFill>"u8;

                public static class Blip
                {
                    public static ReadOnlySpan<byte> GetPrefix()
                        => "<a:blip r:embed=\""u8;

                    public static ReadOnlySpan<byte> GetPostfix()
                        => "\" cstate=\"print\"/>"u8;
                }

                public static class Stretch
                {
                    public static ReadOnlySpan<byte> GetFillRect()
                        => "<a:stretch><a:fillRect/></a:stretch>"u8;
                }
            }

            public static class ShapeProperties
            {
                public static ReadOnlySpan<byte> GetPrefix()
                    => "<xdr:spPr>"u8;

                public static ReadOnlySpan<byte> GetPostfix()
                    => "</xdr:spPr>"u8;

                public static class PresetGeometry
                {
                    public static ReadOnlySpan<byte> GetRect()
                        => "<a:prstGeom prst=\"rect\"/>"u8;
                }
            }

            public static class NonVisualProperties
            {
                public static ReadOnlySpan<byte> GetPrefix()
                    => "<xdr:nvPicPr>"u8;

                public static ReadOnlySpan<byte> GetPostfix()
                    => "</xdr:nvPicPr>"u8;

                public static class Properties
                {
                    public static ReadOnlySpan<byte> GetPrefix()
                        => "<xdr:cNvPr "u8;

                    public static ReadOnlySpan<byte> GetPostfix()
                        => "/>"u8;

                    public static class Id
                    {
                        public static ReadOnlySpan<byte> GetPrefix()
                            => " id=\""u8;

                        public static ReadOnlySpan<byte> GetPostfix()
                            => "\""u8;
                    }

                    public static class Name
                    {
                        public static ReadOnlySpan<byte> GetPrefix()
                            => " name=\""u8;

                        public static ReadOnlySpan<byte> GetPostfix()
                            => "\""u8;
                    }
                }

                public static class PictureProperties
                {
                    public static ReadOnlySpan<byte> GetBody()
                        => "<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr>"u8;
                }
            }
        }

        public static class TwoCellAnchor
        {
            public static ReadOnlySpan<byte> GetPrefix()
                => "<xdr:twoCellAnchor editAs=\"oneCell\">"u8;

            public static ReadOnlySpan<byte> GetPostfix()
                => "</xdr:twoCellAnchor>"u8;
        }

        public static class OneCellAnchor
        {
            public static ReadOnlySpan<byte> GetPrefix()
                => "<xdr:oneCellAnchor>"u8;

            public static ReadOnlySpan<byte> GetPostfix()
                => "</xdr:oneCellAnchor>"u8;

            public static class Size
            {
                public static ReadOnlySpan<byte> GetPrefix()
                    => "<xdr:ext "u8;

                public static ReadOnlySpan<byte> GetPostfix()
                    => "/>"u8;

                public static class Height
                {
                    public static ReadOnlySpan<byte> GetPrefix()
                        => " cx=\""u8;

                    public static ReadOnlySpan<byte> GetPostfix()
                        => "\""u8;
                }

                public static class Width
                {
                    public static ReadOnlySpan<byte> GetPrefix()
                        => " cy=\""u8;

                    public static ReadOnlySpan<byte> GetPostfix()
                        => "\""u8;
                }
            }
        }

        public static class AnchorFrom
        {
            public static ReadOnlySpan<byte> GetPrefix()
                => "<xdr:from>"u8;

            public static ReadOnlySpan<byte> GetPostfix()
                => "</xdr:from>"u8;
        }

        public static class AnchorTo
        {
            public static ReadOnlySpan<byte> GetPrefix()
                => "<xdr:to>"u8;

            public static ReadOnlySpan<byte> GetPostfix()
                => "</xdr:to>"u8;
        }

        public static class AnchorCell
        {
            public static class Column
            {
                public static ReadOnlySpan<byte> GetPrefix()
                    => "<xdr:col>"u8;

                public static ReadOnlySpan<byte> GetPostfix()
                    => "</xdr:col>"u8;
            }

            public static class Row
            {
                public static ReadOnlySpan<byte> GetPrefix()
                    => "<xdr:row>"u8;

                public static ReadOnlySpan<byte> GetPostfix()
                    => "</xdr:row>"u8;
            }

            public static class ColumnOffset
            {
                public static ReadOnlySpan<byte> GetPrefix()
                    => "<xdr:colOff>"u8;

                public static ReadOnlySpan<byte> GetPostfix()
                    => "</xdr:colOff>"u8;
            }

            public static class RowOffset
            {
                public static ReadOnlySpan<byte> GetPrefix()
                    => "<xdr:rowOff>"u8;

                public static ReadOnlySpan<byte> GetPostfix()
                    => "</xdr:rowOff>"u8;
            }
        }
    }
}
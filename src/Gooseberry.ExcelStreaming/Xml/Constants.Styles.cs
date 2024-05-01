// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal static partial class Constants
{
    public static class Styles
    {
        public static ReadOnlySpan<byte> Prefix => "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8;

        public static ReadOnlySpan<byte> Postfix => "</styleSheet>"u8;

        public static class NumberFormats
        {
            public static ReadOnlySpan<byte> Prefix => "<numFmts>"u8;

            public static ReadOnlySpan<byte> Postfix => "</numFmts>"u8;

            public static class Item
            {
                public static ReadOnlySpan<byte> Prefix => "<numFmt numFmtId=\""u8;

                public static ReadOnlySpan<byte> Middle => "\" formatCode=\""u8;

                public static ReadOnlySpan<byte> Postfix => "\"/>"u8;
            }
        }

        public static class Fonts
        {
            public static ReadOnlySpan<byte> Prefix => "<fonts>"u8;

            public static ReadOnlySpan<byte> Postfix => "</fonts>"u8;

            public static class Item
            {
                public static ReadOnlySpan<byte> Prefix => "<font>"u8;

                public static ReadOnlySpan<byte> Postfix => "</font>"u8;

                public static class Size
                {
                    public static ReadOnlySpan<byte> Prefix => "<sz val=\""u8;

                    public static ReadOnlySpan<byte> Postfix => "\"/>"u8;
                }

                public static class Color
                {
                    public static ReadOnlySpan<byte> Prefix => "<color rgb=\""u8;

                    public static ReadOnlySpan<byte> Postfix => "\"/>"u8;
                }

                public static class Name
                {
                    public static ReadOnlySpan<byte> Prefix => "<name val=\""u8;

                    public static ReadOnlySpan<byte> Postfix => "\"/>"u8;
                }

                public static ReadOnlySpan<byte> Bold => "<b val=\"1\"/>"u8;

                public static ReadOnlySpan<byte> Italic => "<i val=\"1\"/>"u8;

                public static ReadOnlySpan<byte> Strike => "<strike val=\"1\"/>"u8;

                public static class Underline
                {
                    public static ReadOnlySpan<byte> Prefix => "<u val=\""u8;

                    public static ReadOnlySpan<byte> Postfix => "\"/>"u8;
                }
            }
        }

        public static class Fills
        {
            public static ReadOnlySpan<byte> Prefix => "<fills>"u8;

            public static ReadOnlySpan<byte> Postfix => "</fills>"u8;

            public static class Item
            {
                public static ReadOnlySpan<byte> Prefix => "<fill>"u8;

                public static ReadOnlySpan<byte> Postfix => "</fill>"u8;

                public static class Pattern
                {
                    public static ReadOnlySpan<byte> Prefix => "<patternFill patternType=\""u8;

                    public static ReadOnlySpan<byte> Medium => "\">"u8;

                    public static ReadOnlySpan<byte> Postfix => "</patternFill>"u8;

                    public static class Color
                    {
                        public static ReadOnlySpan<byte> Prefix => "<fgColor rgb=\""u8;

                        public static ReadOnlySpan<byte> Postfix => "\"/><bgColor auto=\"1\"/>"u8;
                    }
                }
            }
        }

        public static class Borders
        {
            public static ReadOnlySpan<byte> Prefix => "<borders>"u8;

            public static ReadOnlySpan<byte> Postfix => "</borders>"u8;

            public static class Border
            {
                public static ReadOnlySpan<byte> Prefix => "<border>"u8;

                public static ReadOnlySpan<byte> Postfix => "</border>"u8;
            }

            public static class Style
            {
                public static ReadOnlySpan<byte> Prefix => " style=\""u8;

                public static ReadOnlySpan<byte> Postfix => "\""u8;
            }

            public static class Color
            {
                public static ReadOnlySpan<byte> Prefix => "<color rgb=\""u8;

                public static ReadOnlySpan<byte> Postfix => "\"/>"u8;
            }

            public static class Left
            {
                public static ReadOnlySpan<byte> Prefix => "<left"u8;

                public static ReadOnlySpan<byte> Middle => ">"u8;

                public static ReadOnlySpan<byte> Postfix => "</left>"u8;

                public static ReadOnlySpan<byte> Empty => "<left/>"u8;
            }

            public static class Right
            {
                public static ReadOnlySpan<byte> Prefix => "<right"u8;

                public static ReadOnlySpan<byte> Middle => ">"u8;

                public static ReadOnlySpan<byte> Postfix => "</right>"u8;

                public static ReadOnlySpan<byte> Empty => "<right/>"u8;
            }

            public static class Top
            {
                public static ReadOnlySpan<byte> Prefix => "<top"u8;

                public static ReadOnlySpan<byte> Middle => ">"u8;

                public static ReadOnlySpan<byte> Postfix => "</top>"u8;

                public static ReadOnlySpan<byte> Empty => "<top/>"u8;
            }

            public static class Bottom
            {
                public static ReadOnlySpan<byte> Prefix => "<bottom"u8;

                public static ReadOnlySpan<byte> Middle => ">"u8;

                public static ReadOnlySpan<byte> Postfix => "</bottom>"u8;

                public static ReadOnlySpan<byte> Empty => "<bottom/>"u8;
            }
        }

        public static class Colors
        {
            public static ReadOnlySpan<byte> Prefix => "<colors><indexedColors>"u8;

            public static ReadOnlySpan<byte> Postfix => "</indexedColors></colors>"u8;

            public static class Item
            {
                public static ReadOnlySpan<byte> Prefix => "<rgbColor rgb=\""u8;

                public static ReadOnlySpan<byte> Postfix => "\"/>"u8;
            }
        }

        public static class CellStyles
        {
            public static ReadOnlySpan<byte> Prefix => "<cellXfs>"u8;

            public static ReadOnlySpan<byte> Postfix => "</cellXfs>"u8;

            public static class Item
            {
                public static class Open
                {
                    public static ReadOnlySpan<byte> Prefix => "<xf"u8;

                    public static ReadOnlySpan<byte> Postfix => ">"u8;

                    public static class NumberFormatId
                    {
                        public static ReadOnlySpan<byte> Prefix => " numFmtId=\""u8;

                        public static ReadOnlySpan<byte> Postfix => "\" applyNumberFormat=\"1\""u8;

                        public static ReadOnlySpan<byte> Empty => " numFmtId=\"0\" applyNumberFormat=\"0\""u8;
                    }

                    public static class FillId
                    {
                        public static ReadOnlySpan<byte> Prefix => " fillId=\""u8;

                        public static ReadOnlySpan<byte> Postfix => "\" applyFill=\"1\""u8;

                        public static ReadOnlySpan<byte> Empty => " fillId=\"0\" applyFill=\"0\""u8;
                    }

                    public static class FontId
                    {
                        public static ReadOnlySpan<byte> Prefix => " fontId=\""u8;

                        public static ReadOnlySpan<byte> Postfix => "\" applyFont=\"1\""u8;

                        public static ReadOnlySpan<byte> Empty => " fontId=\"0\" applyFont=\"0\""u8;
                    }

                    public static class BorderId
                    {
                        public static ReadOnlySpan<byte> Prefix => " borderId=\""u8;

                        public static ReadOnlySpan<byte> Postfix => "\" applyBorder=\"1\""u8;

                        public static ReadOnlySpan<byte> Empty => " borderId=\"0\" applyBorder=\"0\""u8;
                    }

                    public static ReadOnlySpan<byte> ApplyAlignment => " applyAlignment=\"1\""u8;

                    public static class Alignment
                    {
                        public static ReadOnlySpan<byte> Prefix => "<alignment"u8;

                        public static ReadOnlySpan<byte> Postfix => "/>"u8;

                        public static class Horizontal
                        {
                            public static ReadOnlySpan<byte> Prefix => " horizontal=\""u8;

                            public static ReadOnlySpan<byte> Postfix => "\""u8;
                        }

                        public static class Vertical
                        {
                            public static ReadOnlySpan<byte> Prefix => " vertical=\""u8;

                            public static ReadOnlySpan<byte> Postfix => "\""u8;
                        }

                        public static ReadOnlySpan<byte> WrapText => " wrapText=\"1\""u8;
                    }
                }

                public static ReadOnlySpan<byte> Postfix => "</xf>"u8;
            }
        }
    }
}
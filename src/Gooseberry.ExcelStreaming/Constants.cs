using System.Text;

namespace Gooseberry.ExcelStreaming
{
    internal static class Constants
    {
        public const int DefaultBufferSize = 100 * 1024;

        public const double DefaultBufferFlushThreshold = 0.9;

        public static readonly byte[] XmlPrefix = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>");

        public static class Workbook
        {
            public static readonly byte[] Prefix = Encoding.UTF8.GetBytes($"<workbook xmlns=\"{MainNamespace}\"><sheets>");

            public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</sheets></workbook>");

            public static class Sheet
            {
                public static readonly byte[] StartPrefix = Encoding.UTF8.GetBytes("<sheet name=\"");

                public static readonly byte[] EndPrefix = Encoding.UTF8.GetBytes("\" sheetId=\"");

                public static readonly byte[] EndPostfix = Encoding.UTF8.GetBytes("\" r:id=\"");

                public static readonly byte[] Postfix =
                    Encoding.UTF8.GetBytes("\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>");
            }
        }

        public static class Worksheet
        {
            public static readonly byte[] Prefix =
                Encoding.UTF8.GetBytes($"<worksheet xmlns=\"{MainNamespace}\">");

            public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</worksheet>");

            public static class Columns
            {
                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<cols>");

                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</cols>");

                public static class Item
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<col");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes(" customWidth=\"1\"/>");

                    public static class Min
                    {
                        public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" min=\"");

                        public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"");
                    }

                    public static class Max
                    {
                        public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" max=\"");

                        public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"");
                    }

                    public static class Width
                    {
                        public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" width=\"");

                        public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"");
                    }

                }
            }

            public static class Merges
            {
                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<mergeCells>");

                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</mergeCells>");

                public static class Merge
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<mergeCell ref=\"");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"/>");
                }
            }

            public static class View
            {
                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<sheetViews><sheetView workbookViewId=\"0\"><pane");

                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes(" activePane=\"bottomRight\" state=\"frozen\"/></sheetView></sheetViews>");

                public static class TopLeftCell
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" topLeftCell=\"");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"");
                }
                
                public static class YSplit
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" ySplit=\"");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"");
                }
                
                public static class XSplit
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" xSplit=\"");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"");
                }
            }
            
            public static class SheetData
            {
                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes($"<sheetData>");

                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes($"</sheetData>");

                public static class Row
                {
                    public static class Open
                    {
                        public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<row");

                        public static readonly byte[] Postfix = Encoding.UTF8.GetBytes(">");

                        public static class Height
                        {
                            public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" ht=\"");

                            public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\" customHeight=\"1\"");
                        }
                    }

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</row>");

                    public static class Cell
                    {
                        public static readonly byte[] StringDataType = Encoding.UTF8.GetBytes(" t=\"str\"");

                        public static readonly byte[] NumberDataType = Encoding.UTF8.GetBytes(" t=\"n\"");

                        public static readonly byte[] DateTimeDataType = Encoding.UTF8.GetBytes("");

                        public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<c");

                        public static readonly byte[] Middle = Encoding.UTF8.GetBytes("><v>");

                        public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</v></c>");

                        public static class Style
                        {
                            public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" s=\"");

                            public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"");
                        }

                        public static readonly byte[] Empty = Encoding.UTF8.GetBytes("<c t=\"str\"><v></v></c>");
                    }
                }
            }
        }

        public static class ContentTypes
        {
            public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />" +
                "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");

            public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</Types>");

            public static class Sheet
            {
                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<Override PartName=\"/xl/worksheets/");

                public static readonly byte[] Postfix =
                    Encoding.UTF8.GetBytes(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            }
        }

        public static class WorkbookRelationships
        {
            public static readonly byte[] Prefix =
                Encoding.UTF8.GetBytes("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");

            public static readonly byte[] Postfix =
                Encoding.UTF8.GetBytes("<Relationship Id=\"styles1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>");

            public static class Sheet
            {
                public static readonly byte[] Prefix =
                    Encoding.UTF8.GetBytes("<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"/xl/worksheets/");

                public static readonly byte[] Middle = Encoding.UTF8.GetBytes(".xml\" Id=\"");
                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"/>");
            }
        }

        public static readonly byte[] Relationships = Encoding.UTF8.GetBytes("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
             "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"/xl/workbook.xml\" Id=\"R2196c6c3552b4024\" />" +
             "</Relationships>");

        public static class Styles
        {
            public static readonly byte[] Prefix = Encoding.UTF8.GetBytes($"<styleSheet xmlns=\"{MainNamespace}\">");

            public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</styleSheet>");

            public static class NumberFormats
            {
                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<numFmts>");

                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</numFmts>");

                public static class Item
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<numFmt numFmtId=\"");

                    public static readonly byte[] Middle = Encoding.UTF8.GetBytes("\" formatCode=\"");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"/>");
                }
            }

            public static class Fonts
            {
                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<fonts>");

                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</fonts>");

                public static class Item
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<font>");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</font>");

                    public static class Size
                    {
                        public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<sz val=\"");

                        public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"/>");
                    }

                    public static class Color
                    {
                        public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<color rgb=\"");

                        public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"/>");
                    }

                    public static class Name
                    {
                        public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<name val=\"");

                        public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"/>");
                    }

                    public static readonly byte[] Bold = Encoding.UTF8.GetBytes("<b val=\"1\"/>");
                    
                    public static readonly byte[] Italic = Encoding.UTF8.GetBytes("<i val=\"1\"/>");
                    
                    public static readonly byte[] Strike = Encoding.UTF8.GetBytes("<strike val=\"1\"/>");
                    
                    public static class Underline
                    {
                        public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<u val=\"");

                        public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"/>");
                    }
                }
            }

            public static class Fills
            {
                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<fills>");

                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</fills>");

                public static class Item
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<fill>");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</fill>");

                    public static class Pattern
                    {
                        public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<patternFill patternType=\"");

                        public static readonly byte[] Medium = Encoding.UTF8.GetBytes("\">");

                        public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</patternFill>");

                        public static class Color
                        {
                            public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<fgColor rgb=\"");

                            public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"/><bgColor auto=\"1\"/>");
                        }
                    }
                }
            }

            public static class Borders
            {
                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<borders>");

                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</borders>");

                public static class Border
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<border>");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</border>");
                }

                public static class Style
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" style=\"");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"");
                }

                public static class Color
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<color rgb=\"");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"/>");
                }

                public static class Left
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<left");

                    public static readonly byte[] Middle = Encoding.UTF8.GetBytes(">");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</left>");

                    public static readonly byte[] Empty = Encoding.UTF8.GetBytes("<left/>");
                }

                public static class Right
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<right");

                    public static readonly byte[] Middle = Encoding.UTF8.GetBytes(">");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</right>");

                    public static readonly byte[] Empty = Encoding.UTF8.GetBytes("<right/>");
                }

                public static class Top
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<top");

                    public static readonly byte[] Middle = Encoding.UTF8.GetBytes(">");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</top>");

                    public static readonly byte[] Empty = Encoding.UTF8.GetBytes("<top/>");
                }

                public static class Bottom
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<bottom");

                    public static readonly byte[] Middle = Encoding.UTF8.GetBytes(">");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</bottom>");

                    public static readonly byte[] Empty = Encoding.UTF8.GetBytes("<bottom/>");
                }
            }

            public static class Colors
            {
                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<colors><indexedColors>");

                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</indexedColors></colors>");

                public static class Item
                {
                    public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<rgbColor rgb=\"");

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"/>");
                }
            }

            public static class CellStyles
            {
                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<cellXfs>");

                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</cellXfs>");

                public static class Item
                {
                    public static class Open
                    {
                        public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<xf");

                        public static readonly byte[] Postfix = Encoding.UTF8.GetBytes(">");

                        public static class NumberFormatId
                        {
                            public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" numFmtId=\"");

                            public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\" applyNumberFormat=\"1\"");

                            public static readonly byte[] Empty = Encoding.UTF8.GetBytes(" numFmtId=\"0\" applyNumberFormat=\"0\"");
                        }

                        public static class FillId
                        {
                            public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" fillId=\"");

                            public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\" applyFill=\"1\"");

                            public static readonly byte[] Empty = Encoding.UTF8.GetBytes(" fillId=\"0\" applyFill=\"0\"");
                        }

                        public static class FontId
                        {
                            public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" fontId=\"");

                            public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\" applyFont=\"1\"");

                            public static readonly byte[] Empty = Encoding.UTF8.GetBytes(" fontId=\"0\" applyFont=\"0\"");
                        }

                        public static class BorderId
                        {
                            public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" borderId=\"");

                            public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\" applyBorder=\"1\"");

                            public static readonly byte[] Empty = Encoding.UTF8.GetBytes(" borderId=\"0\" applyBorder=\"0\"");
                        }

                        public static readonly byte[] ApplyAlignment = Encoding.UTF8.GetBytes(" applyAlignment=\"1\"");

                        public static class Alignment
                        {
                            public static readonly byte[] Prefix = Encoding.UTF8.GetBytes("<alignment");

                            public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("/>");

                            public static class Horizontal
                            {
                                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" horizontal=\"");

                                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"");
                            }

                            public static class Vertical
                            {
                                public static readonly byte[] Prefix = Encoding.UTF8.GetBytes(" vertical=\"");

                                public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("\"");
                            }

                            public static readonly byte[] WrapText = Encoding.UTF8.GetBytes(" wrapText=\"1\"");
                        }
                    }

                    public static readonly byte[] Postfix = Encoding.UTF8.GetBytes("</xf>");
                }
            }
        }

        private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    }
}
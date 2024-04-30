using System.Drawing;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Tests.Converters;
using Gooseberry.ExcelStreaming.Tests.Extensions;
using Alignment = Gooseberry.ExcelStreaming.Styles.Alignment;
using Border = Gooseberry.ExcelStreaming.Styles.Border;
using Borders = Gooseberry.ExcelStreaming.Styles.Borders;
using Color = Gooseberry.ExcelStreaming.Styles.Color;
using Fill = Gooseberry.ExcelStreaming.Styles.Fill;
using Font = Gooseberry.ExcelStreaming.Styles.Font;
using MarkerType = DocumentFormat.OpenXml.Drawing.Spreadsheet.MarkerType;
using Point = System.Drawing.Point;

namespace Gooseberry.ExcelStreaming.Tests.Excel;

public static class ExcelReader
{
    public static IReadOnlyCollection<Style> ReadStyles(Stream stream)
    {
        using var spreadsheet = SpreadsheetDocument.Open(stream, isEditable: false);

        return GetStyles(spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet)
            .ToArray();
    }

    public static IReadOnlyCollection<string> ReadSharedStrings(Stream stream)
    {
        using var spreadsheet = SpreadsheetDocument.Open(stream, isEditable: false);

        return spreadsheet.WorkbookPart!.SharedStringTablePart!.SharedStringTable
            !.OfType<SharedStringItem>()
            .Select(i => i.Text!.Text)
            .ToArray();
    }

    private static Dictionary<CellValues, CellValueType> CellTypesMap = new()
    {
        { CellValues.String, CellValueType.String },
        { CellValues.Number, CellValueType.Number },
        { CellValues.Date, CellValueType.DateTime },
        { CellValues.SharedString, CellValueType.SharedString }
    };


    private static Dictionary<PatternValues, FillPattern> PatternMap = new()
    {
        { PatternValues.None, FillPattern.None },
        { PatternValues.Gray125, FillPattern.Gray125 },
        { PatternValues.Solid, FillPattern.Solid }
    };

    private static Dictionary<UnderlineValues, Styles.Underline> UnderlineMap = new()
    {
        { UnderlineValues.None, ExcelStreaming.Styles.Underline.None },
        { UnderlineValues.Single, ExcelStreaming.Styles.Underline.Single },
        { UnderlineValues.Double, ExcelStreaming.Styles.Underline.Double },
        { UnderlineValues.SingleAccounting, ExcelStreaming.Styles.Underline.SingleAccounting },
        { UnderlineValues.DoubleAccounting, ExcelStreaming.Styles.Underline.DoubleAccounting }
    };

    private static Dictionary<HorizontalAlignmentValues, HorizontalAlignment> HorizontalMap = new()
    {
        { HorizontalAlignmentValues.Center, HorizontalAlignment.Center },
        { HorizontalAlignmentValues.CenterContinuous, HorizontalAlignment.CenterContinuous },
        { HorizontalAlignmentValues.Distributed, HorizontalAlignment.Distributed },
        { HorizontalAlignmentValues.Fill, HorizontalAlignment.Fill },
        { HorizontalAlignmentValues.General, HorizontalAlignment.General },
        { HorizontalAlignmentValues.Justify, HorizontalAlignment.Justify },
        { HorizontalAlignmentValues.Left, HorizontalAlignment.Left },
        { HorizontalAlignmentValues.Right, HorizontalAlignment.Right }
    };

    private static Dictionary<VerticalAlignmentValues, VerticalAlignment> VerticalMap = new()
    {
        { VerticalAlignmentValues.Bottom, VerticalAlignment.Bottom },
        { VerticalAlignmentValues.Center, VerticalAlignment.Center },
        { VerticalAlignmentValues.Distributed, VerticalAlignment.Distributed },
        { VerticalAlignmentValues.Justify, VerticalAlignment.Justify },
        { VerticalAlignmentValues.Top, VerticalAlignment.Top }
    };


    public static IReadOnlyCollection<Sheet> ReadSheets(Stream stream)
    {
        using var spreadsheet = SpreadsheetDocument.Open(stream, isEditable: false);

        var styles = GetStyles(spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet)
            .Select((style, index) => (index, style))
            .ToDictionary(x => (uint)x.index, x => x.style);

        var sheets = spreadsheet
            .WorkbookPart
            !.Workbook
            .Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .Select(s => (Id: s.Id!.Value!, Name: s.Name!.Value!))
            .ToDictionary(s => s.Id, s => s.Name);

        return spreadsheet.WorkbookPart!.Parts!
            .Where(part => part.OpenXmlPart is WorksheetPart)
            .Select(
                part =>
                    new Sheet(
                        Name: sheets[part.RelationshipId],
                        GetRows((WorksheetPart)part.OpenXmlPart),
                        GetColumns((WorksheetPart)part.OpenXmlPart),
                        GetMerges((WorksheetPart)part.OpenXmlPart),
                        GetPictures((WorksheetPart)part.OpenXmlPart)))
            .ToArray();

        IReadOnlyCollection<Picture> GetPictures(WorksheetPart sheetPart)
        {
            var drawingsPart = sheetPart.DrawingsPart;

            if (drawingsPart is null)
                return Array.Empty<Picture>();

            var oneCellAnchorPictures = drawingsPart.RootElement!.Descendants<OneCellAnchor>()
                .Select(v => GetOneCellAnchorPicture(drawingsPart, v));

            var twoCellAnchorPictures = drawingsPart.RootElement!.Descendants<TwoCellAnchor>()
                .Select(v => GetTwoCellAnchorPicture(drawingsPart, v));

            return oneCellAnchorPictures.Union(twoCellAnchorPictures).ToArray();
        }

        Picture GetOneCellAnchorPicture(OpenXmlPartContainer drawingsPart, OneCellAnchor oneCellAnchor)
        {
            var placement = new PicturePlacement(
                GetAnchorCell(oneCellAnchor.FromMarker!),
                GetSize(oneCellAnchor.Extent!));

            var imagePart = (ImagePart)drawingsPart.GetPartById(oneCellAnchor.Descendants<Blip>().Single().Embed?.Value ?? "");
            var format = ContentTypeConverter.ToPictureFormat(imagePart.ContentType);

            return new Picture(imagePart.GetStream().ToArray(), placement, format);
        }

        Picture GetTwoCellAnchorPicture(OpenXmlPartContainer drawingsPart, TwoCellAnchor twoCellAnchor)
        {
            var placement = new PicturePlacement(
                GetAnchorCell(twoCellAnchor.FromMarker!),
                GetAnchorCell(twoCellAnchor.ToMarker!));

            var imagePart = (ImagePart)drawingsPart.GetPartById(twoCellAnchor.Descendants<Blip>().Single().Embed?.Value ?? "");
            var format = ContentTypeConverter.ToPictureFormat(imagePart.ContentType);

            return new Picture(imagePart.GetStream().ToArray(), placement, format);
        }

        Size GetSize(Extent extent)
            => new(width: (int)extent.Cy!.Value, height: (int)extent.Cx!.Value);

        AnchorCell GetAnchorCell(MarkerType marker)
        {
            return new AnchorCell(
                Row: int.Parse(marker.RowId?.Text ?? ""),
                Column: int.Parse(marker.ColumnId?.Text ?? ""),
                Offset: new Point(
                    x: int.Parse(marker.ColumnOffset?.Text ?? ""),
                    y: int.Parse(marker.RowOffset?.Text ?? "")));
        }

        // sheetPart.DrawingsPart.RootElement.Descendants<OneCellAnchor>().First().Descendants<Picture>().First()
        IReadOnlyCollection<Column> GetColumns(WorksheetPart sheetPart)
        {
            return sheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Column>()
                .Select(column => new Column((decimal)column.Width!.Value))
                .ToArray();
        }

        IReadOnlyCollection<Row> GetRows(WorksheetPart sheetPart)
        {
            return sheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>()
                .Select(data => new Excel.Row(GetCells(data), (decimal?)data.Height?.Value))
                .ToArray();
        }

        IReadOnlyCollection<string> GetMerges(WorksheetPart sheetPart)
        {
            return sheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.MergeCell>()
                .Select(data => data.Reference!.ToString()!)
                .ToArray();
        }

        IReadOnlyCollection<Cell> GetCells(DocumentFormat.OpenXml.Spreadsheet.Row row)
            => row.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Select(GetCell).ToArray();


        Cell GetCell(DocumentFormat.OpenXml.Spreadsheet.Cell cell)
        {
            CellValueType? valueType = cell.DataType?.Value == null
                ? null
                : CellTypesMap.TryGetValue(cell.DataType.Value, out var value)
                    ? value
                    : throw new InvalidCastException($"Cannot convert {cell.DataType?.Value} to CellValueType.");


            Style? style = null;

            if (cell!.StyleIndex?.HasValue == true)
                if (styles.TryGetValue(cell.StyleIndex!.Value!, out var styleValue))
                    style = styleValue;

            return new Cell(cell.CellValue!.Text, valueType, style);
        }
    }

    private static IEnumerable<Style> GetStyles(Stylesheet styles)
    {
        var numberFormats = styles.NumberingFormats
            !.OfType<NumberingFormat>()
            .ToDictionary(n => n.NumberFormatId!, n => n.FormatCode!.Value!);

        var fills = styles.Fills!.OfType<DocumentFormat.OpenXml.Spreadsheet.Fill>()
            .Select(GetFill)
            .ToArray();

        var borders = styles.Borders!.OfType<DocumentFormat.OpenXml.Spreadsheet.Border>()
            .Select(GetBorders)
            .ToArray();

        var fonts = styles.Fonts!.OfType<DocumentFormat.OpenXml.Spreadsheet.Font>()
            .Select(GetFont)
            .ToArray();

        return styles.CellFormats
            !.OfType<CellFormat>()
            .Select(
                s =>
                    new Style(
                        Format: numberFormats[s.NumberFormatId!],
                        Fill: s.FillId?.HasValue == true ? fills[(int)s.FillId.Value] : null,
                        Borders: s.BorderId?.HasValue == true ? borders[(int)s.BorderId.Value] : null,
                        Font: s.FontId?.HasValue == true ? fonts[(int)s.FontId.Value] : null,
                        Alignment: s.Alignment != null ? GetAlignment(s.Alignment) : null));
    }

    private static Fill GetFill(DocumentFormat.OpenXml.Spreadsheet.Fill fill)
    {
        if (fill.PatternFill == null)
            throw new InvalidOperationException("PatternFill should not be empty.");

        var pattern = fill.PatternFill.PatternType?.Value != null &&
            PatternMap.TryGetValue(fill.PatternFill.PatternType.Value, out var value)
                ? value
                : throw new InvalidCastException($"Cannot convert {fill.PatternFill.PatternType?.Value} to FillPattern.");

        var color = GetColor(fill.PatternFill.ForegroundColor?.Rgb);

        return new Fill(color, pattern);
    }

    private static Borders GetBorders(DocumentFormat.OpenXml.Spreadsheet.Border border)
    {
        return new Borders(
            left: GetBorder(border.LeftBorder),
            right: GetBorder(border.RightBorder),
            top: GetBorder(border.TopBorder),
            bottom: GetBorder(border.BottomBorder));
    }

    private static Border? GetBorder(BorderPropertiesType? border)
    {
        if (border == null || border.Style == null && border.Color == null)
            return null;

        var style = border.Style?.Value.AsString() switch
        {
            "thin" => BorderStyle.Thin,
            _ => throw new InvalidCastException($"Cannot convert {border.Style?.Value} to BorderStyle.")
        };

        var color = GetColor(border.Color?.Rgb);

        if (!color.HasValue)
            throw new InvalidOperationException("Color should be set for border.");

        return new Border(style, color.Value);
    }

    private static Font GetFont(DocumentFormat.OpenXml.Spreadsheet.Font font)
    {
        var size = (int)(font.FontSize?.Val?.Value ?? 0);
        var name = font.FontName?.Val?.Value;
        var color = GetColor(font.Color?.Rgb);
        var bold = font.Bold?.Val?.Value ?? false;
        var italic = font.Italic?.Val?.Value ?? false;
        var strike = font.Strike?.Val?.Value ?? false;

        var underline = font.Underline?.Val?.Value != null &&
            UnderlineMap.TryGetValue(font.Underline.Val.Value, out var value)
                ? value
                : throw new InvalidCastException($"Cannot convert {font.Underline?.Val?.Value} to Underline.");

        return new Font(size, name, color, bold, italic, strike, underline);
    }

    private static Alignment GetAlignment(DocumentFormat.OpenXml.Spreadsheet.Alignment alignment)
    {
        HorizontalAlignment? horizontal = alignment.Horizontal?.Value == null
            ? null
            : HorizontalMap.TryGetValue(alignment.Horizontal.Value, out var hValue)
                ? hValue
                : throw new InvalidCastException($"Cannot convert {alignment.Horizontal?.Value} to HorizontalAlignment.");

        VerticalAlignment? vertical = alignment.Vertical?.Value == null
            ? null
            : VerticalMap.TryGetValue(alignment.Vertical.Value, out var vValue)
                ? vValue
                : throw new InvalidCastException($"Cannot convert {alignment.Horizontal?.Value} to VerticalAlignment.");

        var wrapText = alignment.WrapText?.Value == true;

        return new Alignment(horizontal, vertical, wrapText);
    }

    private static Color? GetColor(HexBinaryValue? rgbColor)
    {
        var rgbValue = rgbColor?.Value;

        if (!string.IsNullOrEmpty(rgbValue))
            return new Color(uint.Parse(rgbValue, NumberStyles.HexNumber));

        return null;
    }

    public static string AsString(this IEnumValue value)
        => value.Value;
}
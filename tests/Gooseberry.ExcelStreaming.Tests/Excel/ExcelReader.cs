using System.Drawing;
using System.Globalization;
using System.Text;
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
using Color = System.Drawing.Color;
using Fill = Gooseberry.ExcelStreaming.Styles.Fill;
using Font = Gooseberry.ExcelStreaming.Styles.Font;
using Format = Gooseberry.ExcelStreaming.Styles.Format;
using MarkerType = DocumentFormat.OpenXml.Drawing.Spreadsheet.MarkerType;
using Point = System.Drawing.Point;
using Underline = Gooseberry.ExcelStreaming.Styles.Underline;

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
                Row: uint.Parse(marker.RowId?.Text ?? ""),
                Column: uint.Parse(marker.ColumnId?.Text ?? ""),
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
                .Select(data => new Excel.Row(
                    GetCells(data),
                    (decimal?)data.Height?.Value,
                    data.OutlineLevel?.Value,
                    data.Hidden?.Value,
                    data.Collapsed?.Value))
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

            return new Cell(cell.CellValue?.Text ?? string.Empty, valueType, style);
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
                        Format: numberFormats.TryGetValue(s.NumberFormatId!, out var numberFormat) ? numberFormat : new Format((int)s.NumberFormatId!.Value),
                        Fill: s.FillId?.HasValue == true ? fills[(int)s.FillId.Value] : null,
                        Borders: s.BorderId?.HasValue == true ? borders[(int)s.BorderId.Value] : null,
                        Font: s.FontId?.HasValue == true ? fonts[(int)s.FontId.Value] : null,
                        Alignment: s.Alignment != null ? GetAlignment(s.Alignment) : null));
    }

    private static Fill GetFill(DocumentFormat.OpenXml.Spreadsheet.Fill fill)
    {
        if (fill.PatternFill == null)
            throw new InvalidOperationException("PatternFill should not be empty.");

        var pattern = new FillPattern(Encoding.UTF8.GetBytes(fill.PatternFill.PatternType?.Value.AsString() ?? "none"));
        var color = GetColor(fill.PatternFill.ForegroundColor?.Rgb);

        return new Fill(color, pattern);
    }

    private static Borders GetBorders(DocumentFormat.OpenXml.Spreadsheet.Border border)
    {
        return new Borders(
            Left: GetBorder(border.LeftBorder),
            Right: GetBorder(border.RightBorder),
            Top: GetBorder(border.TopBorder),
            Bottom: GetBorder(border.BottomBorder));
    }

    private static Border? GetBorder(BorderPropertiesType? border)
    {
        if (border == null || border.Style == null && border.Color == null)
            return null;

        var style = new BorderStyle(Encoding.UTF8.GetBytes(border.Style?.Value.AsString() ?? "none"));

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

        var underline = new Underline(Encoding.UTF8.GetBytes(font.Underline?.Val?.Value.AsString() ?? "none"));


        return new Font(size, name, color, bold, italic, strike, underline);
    }

    private static Alignment GetAlignment(DocumentFormat.OpenXml.Spreadsheet.Alignment alignment)
    {
        HorizontalAlignment? horizontal = alignment.Horizontal?.Value == null
            ? null
            : new HorizontalAlignment(Encoding.UTF8.GetBytes(alignment.Horizontal?.Value.AsString() ?? "general"));

        VerticalAlignment? vertical = alignment.Vertical?.Value == null
            ? null
            : new VerticalAlignment(Encoding.UTF8.GetBytes(alignment.Vertical?.Value.AsString() ?? "top"));

        var wrapText = alignment.WrapText?.Value == true;

        return new Alignment(horizontal, vertical, wrapText);
    }

    private static Color? GetColor(HexBinaryValue? rgbColor)
    {
        var rgbValue = rgbColor?.Value;

        if (!string.IsNullOrEmpty(rgbValue))
            return Color.FromArgb(int.Parse(rgbValue, NumberStyles.HexNumber));

        return null;
    }

    public static string AsString(this IEnumValue value)
        => value.Value;
}
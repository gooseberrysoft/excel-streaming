using System.Drawing;
using Gooseberry.ExcelStreaming.Styles.Records;

namespace Gooseberry.ExcelStreaming.Styles;

public sealed class StylesSheetBuilder
{
    private const string GeneralFormat = "General";

    private readonly Dictionary<string, FormatRecord> _formats = new();
    private readonly List<Fill> _fills = new();
    private readonly List<Font> _fonts = new();
    private readonly List<Borders> _borders = new();
    private readonly List<StyleRecord> _styles = new();

    private readonly StyleReference _defaultDateStyle;
    private readonly StyleReference _defaultHyperlinkStyle;

    internal static readonly StylesSheet Default = new StylesSheetBuilder().Build();

    public StylesSheetBuilder()
    {
        _formats[GeneralFormat] = new FormatRecord(0, GeneralFormat);
        _borders.Add(new Borders());
        _fills.Add(new Fill(FillPattern.None));
        _fills.Add(new Fill(FillPattern.Gray125));

        GetOrAdd(new Style(
            Format: GeneralFormat,
            Font: new Font(Size: 11, Name: null, Color: Color.Black, Bold: false)));

        _defaultDateStyle = GetOrAdd(new Style(Format: StandardFormat.DayMonthYear4WithSlashes));

        _defaultHyperlinkStyle = GetOrAdd(new Style(Font: new Font(Color: Color.Navy, Underline: Underline.Single)));
    }

    public StyleReference GetOrAdd(Style style)
    {
        var formatId = style.Format?.FormatId ??
            (!string.IsNullOrEmpty(style.Format?.CustomFormat) ? GetOrAddFormatId(style.Format.Value.CustomFormat) : null);
        var fillId = style.Fill.HasValue ? GetOrAddFillId(style.Fill.Value) : (int?)null;
        var fontId = style.Font.HasValue ? GetOrAddFontId(style.Font.Value) : (int?)null;
        var bordersId = style.Borders.HasValue ? GetOrAddBordersId(style.Borders.Value) : (int?)null;

        var styleRecord = new StyleRecord(formatId, fillId, fontId, bordersId, style.Alignment);

        var index = _styles.IndexOf(styleRecord);
        if (index >= 0)
            return new StyleReference(index);

        _styles.Add(styleRecord);
        return new StyleReference(_styles.Count - 1);
    }

    public StylesSheet Build()
    {
        using var writer = new StylesWriter();

        writer.AddNumberFormats(_formats.Values);
        writer.AddFonts(_fonts);
        writer.AddFills(_fills);
        writer.AddBorders(_borders);
        writer.AddCellStyles(_styles);

        var stylesData = writer.GetWrittenData();

        return new StylesSheet(stylesData, _defaultDateStyle, _defaultHyperlinkStyle);
    }

    private int GetOrAddFormatId(string format)
    {
        if (!_formats.TryGetValue(format, out var formatRecord))
        {
            // Excel has predefined number formats, so to add new formats we add some shift
            // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
            const int formatIdShift = 1000;
            formatRecord = new FormatRecord(_formats.Count + formatIdShift, format);
            _formats[format] = formatRecord;
        }

        return formatRecord.Id;
    }

    private int GetOrAddFillId(Fill fill)
    {
        var index = _fills.IndexOf(fill);
        if (index >= 0)
            return index;

        _fills.Add(fill);
        return _fills.Count - 1;
    }

    private int GetOrAddFontId(in Font font)
    {
        var index = _fonts.IndexOf(font);
        if (index >= 0)
            return index;

        _fonts.Add(font);
        return _fonts.Count - 1;
    }

    private int GetOrAddBordersId(in Borders borders)
    {
        var index = _borders.IndexOf(borders);
        if (index >= 0)
            return index;

        _borders.Add(borders);
        return _borders.Count - 1;
    }
}
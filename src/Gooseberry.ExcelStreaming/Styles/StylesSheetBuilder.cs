using Gooseberry.ExcelStreaming.Styles.Records;

namespace Gooseberry.ExcelStreaming.Styles;

public sealed class StylesSheetBuilder
{
    private const uint Black = 0xff000000;
    private const uint Navy = 0x0000007B;

    private const string GeneralFormat = "General";
    private const string DefaultDateFormat = "dd.mm.yyyy";

    private readonly Dictionary<string, FormatRecord> _formats = new();
    private readonly List<Fill> _fills = new();
    private readonly List<Font> _fonts = new();
    private readonly List<Borders> _borders = new();
    private readonly List<StyleRecord> _styles = new();

    private readonly StyleReference _generalStyle;
    private readonly StyleReference _defaultDateStyle;
    private readonly StyleReference _defaultHyperlinkStyle;

    internal static readonly StylesSheet Default = new StylesSheetBuilder().Build();

    public StylesSheetBuilder()
    {
        _formats[GeneralFormat] = new FormatRecord(0, GeneralFormat);
        _borders.Add(new Borders());
        _fills.Add(new Fill(pattern: FillPattern.None, color: null));
        _fills.Add(new Fill(pattern: FillPattern.Gray125, color: null));

        _generalStyle = GetOrAdd(new Style(
            format: GeneralFormat,
            font: new Font(size: 11, name: null, color: Black, bold: false)));

        _defaultDateStyle = GetOrAdd(new Style(format: DefaultDateFormat));

        _defaultHyperlinkStyle = GetOrAdd(new Style(font: new Font(color: Navy, underline: Underline.Single)));
    }

    public StyleReference GetOrAdd(Style style)
    {
        var formatId = !string.IsNullOrEmpty(style.Format) ? GetOrAddFormatId(style.Format!) : (int?)null;
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

        return new StylesSheet(stylesData, _generalStyle, _defaultDateStyle, _defaultHyperlinkStyle);
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
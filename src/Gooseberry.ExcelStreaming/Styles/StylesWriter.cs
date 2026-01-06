using System.Text;
using Gooseberry.ExcelStreaming.Styles.Records;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.Styles;

internal sealed class StylesWriter : IDisposable
{
    private readonly BuffersChain _buffer;
    private readonly Encoder _encoder;
    private const string Hex8 = "X8";

    public StylesWriter()
    {
        _buffer = new BuffersChain(bufferMinSize: 16 * 1024);
        _encoder = Encoding.UTF8.GetEncoder();

        Constants.XmlPrefix.WriteTo(_buffer);
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8.WriteTo(_buffer);
    }

    public void AddNumberFormats(IReadOnlyCollection<FormatRecord> formats)
    {
        var span = _buffer.GetSpan();
        var written = 0;

        "<numFmts>"u8.WriteTo(_buffer, ref span, ref written);

        foreach (var format in formats)
        {
            "<numFmt numFmtId=\""u8.WriteTo(_buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(format.Id, _buffer, ref span, ref written);
            "\" formatCode=\""u8.WriteTo(_buffer, ref span, ref written);
            format.Format.WriteEscapedTo(_buffer, _encoder, ref span, ref written);
            "\"/>"u8.WriteTo(_buffer, ref span, ref written);
        }

        "</numFmts>"u8.WriteTo(_buffer, ref span, ref written);

        _buffer.Advance(written);
    }

    public void AddFills(IReadOnlyCollection<Fill> fills)
    {
        var span = _buffer.GetSpan();
        var written = 0;

        "<fills>"u8.WriteTo(_buffer, ref span, ref written);

        foreach (var fill in fills)
        {
            "<fill>"u8.WriteTo(_buffer, ref span, ref written);

            "<patternFill patternType=\""u8.WriteTo(_buffer, ref span, ref written);
            fill.Pattern.Value.WriteTo(_buffer, ref span, ref written);
            "\">"u8.WriteTo(_buffer, ref span, ref written);

            if (fill.Color.HasValue)
            {
                "<fgColor rgb=\""u8.WriteTo(_buffer, ref span, ref written);
                Utf8SpanFormattableWriter.WriteValue(fill.Color.Value.ToArgb(), Hex8, null, _buffer, ref span, ref written);
                "\"/><bgColor auto=\"1\"/>"u8.WriteTo(_buffer, ref span, ref written);
            }

            "</patternFill>"u8.WriteTo(_buffer, ref span, ref written);

            "</fill>"u8.WriteTo(_buffer, ref span, ref written);
        }

        "</fills>"u8.WriteTo(_buffer, ref span, ref written);

        _buffer.Advance(written);
    }

    public void AddCellStyles(IReadOnlyCollection<StyleRecord> styles)
    {
        var span = _buffer.GetSpan();
        var written = 0;

        "<cellXfs>"u8.WriteTo(_buffer, ref span, ref written);

        foreach (var style in styles)
        {
            "<xf"u8.WriteTo(_buffer, ref span, ref written);

            if (style.FormatId.HasValue)
            {
                " numFmtId=\""u8.WriteTo(_buffer, ref span, ref written);
                Utf8SpanFormattableWriter.WriteValue(style.FormatId.Value, _buffer, ref span, ref written);
                "\" applyNumberFormat=\"1\""u8.WriteTo(_buffer, ref span, ref written);
            }
            else
                " numFmtId=\"0\" applyNumberFormat=\"0\""u8.WriteTo(_buffer, ref span, ref written);

            if (style.FillId.HasValue)
            {
                " fillId=\""u8.WriteTo(_buffer, ref span, ref written);
                Utf8SpanFormattableWriter.WriteValue(style.FillId.Value, _buffer, ref span, ref written);
                "\" applyFill=\"1\""u8.WriteTo(_buffer, ref span, ref written);
            }
            else
                " fillId=\"0\" applyFill=\"0\""u8.WriteTo(_buffer, ref span, ref written);

            if (style.FontId.HasValue)
            {
                " fontId=\""u8.WriteTo(_buffer, ref span, ref written);
                Utf8SpanFormattableWriter.WriteValue(style.FontId.Value, _buffer, ref span, ref written);
                "\" applyFont=\"1\""u8.WriteTo(_buffer, ref span, ref written);
            }
            else
                " fontId=\"0\" applyFont=\"0\""u8.WriteTo(_buffer, ref span, ref written);

            if (style.BorderId.HasValue)
            {
                " borderId=\""u8.WriteTo(_buffer, ref span, ref written);
                Utf8SpanFormattableWriter.WriteValue(style.BorderId.Value, _buffer, ref span, ref written);
                "\" applyBorder=\"1\""u8.WriteTo(_buffer, ref span, ref written);
            }
            else
                " borderId=\"0\" applyBorder=\"0\""u8.WriteTo(_buffer, ref span, ref written);

            if (style.Alignment.HasValue)
                " applyAlignment=\"1\""u8.WriteTo(_buffer, ref span, ref written);

            ">"u8.WriteTo(_buffer, ref span, ref written);

            if (style.Alignment.HasValue)
                AddAlignment(style.Alignment.Value, ref span, ref written);

            "</xf>"u8.WriteTo(_buffer, ref span, ref written);
        }

        "</cellXfs>"u8.WriteTo(_buffer, ref span, ref written);

        _buffer.Advance(written);
    }

    private void AddAlignment(Alignment alignment, ref Span<byte> span, ref int written)
    {
        "<alignment"u8.WriteTo(_buffer, ref span, ref written);

        if (alignment.Horizontal.HasValue)
        {
            " horizontal=\""u8.WriteTo(_buffer, ref span, ref written);
            alignment.Horizontal.Value.Value.WriteTo(_buffer, ref span, ref written);
            "\""u8.WriteTo(_buffer, ref span, ref written);
        }

        if (alignment.Vertical.HasValue)
        {
            " vertical=\""u8.WriteTo(_buffer, ref span, ref written);
            alignment.Vertical.Value.Value.WriteTo(_buffer, ref span, ref written);
            "\""u8.WriteTo(_buffer, ref span, ref written);
        }

        if (alignment.WrapText)
            " wrapText=\"1\""u8.WriteTo(_buffer, ref span, ref written);

        "/>"u8.WriteTo(_buffer, ref span, ref written);
    }

    public void AddFonts(IReadOnlyCollection<Font> fonts)
    {
        var span = _buffer.GetSpan();
        var written = 0;

        "<fonts>"u8.WriteTo(_buffer, ref span, ref written);

        foreach (var font in fonts)
        {
            "<font>"u8.WriteTo(_buffer, ref span, ref written);

            "<sz val=\""u8.WriteTo(_buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(font.Size, _buffer, ref span, ref written);
            "\"/>"u8.WriteTo(_buffer, ref span, ref written);

            if (!string.IsNullOrEmpty(font.Name))
            {
                "<name val=\""u8.WriteTo(_buffer, ref span, ref written);
                font.Name.WriteTo(_buffer, _encoder, ref span, ref written);
                "\"/>"u8.WriteTo(_buffer, ref span, ref written);
            }

            if (font.Color.HasValue)
            {
                "<color rgb=\""u8.WriteTo(_buffer, ref span, ref written);
                Utf8SpanFormattableWriter.WriteValue(font.Color.Value.ToArgb(), Hex8, null, _buffer, ref span, ref written);
                "\"/>"u8.WriteTo(_buffer, ref span, ref written);
            }

            if (font.Bold)
                "<b val=\"1\"/>"u8.WriteTo(_buffer, ref span, ref written);

            if (font.Italic)
                "<i val=\"1\"/>"u8.WriteTo(_buffer, ref span, ref written);

            if (font.Strike)
                "<strike val=\"1\"/>"u8.WriteTo(_buffer, ref span, ref written);

            "<u val=\""u8.WriteTo(_buffer, ref span, ref written);
            font.Underline.Value.WriteTo(_buffer, ref span, ref written);
            "\"/>"u8.WriteTo(_buffer, ref span, ref written);

            "</font>"u8.WriteTo(_buffer, ref span, ref written);
        }

        "</fonts>"u8.WriteTo(_buffer, ref span, ref written);

        _buffer.Advance(written);
    }

    public void AddBorders(IReadOnlyCollection<Borders> borders)
    {
        var span = _buffer.GetSpan();
        var written = 0;

        "<borders>"u8.WriteTo(_buffer, ref span, ref written);

        foreach (var border in borders)
        {
            "<border>"u8.WriteTo(_buffer, ref span, ref written);

            AddBorder(
                border.Left,
                "<left/>"u8,
                "<left"u8,
                "</left>"u8,
                ref span,
                ref written);

            AddBorder(
                border.Right,
                "<right/>"u8,
                "<right"u8,
                "</right>"u8,
                ref span,
                ref written);

            AddBorder(
                border.Top,
                "<top/>"u8,
                "<top"u8,
                "</top>"u8,
                ref span,
                ref written);

            AddBorder(
                border.Bottom,
                "<bottom/>"u8,
                "<bottom"u8,
                "</bottom>"u8,
                ref span,
                ref written);

            "</border>"u8.WriteTo(_buffer, ref span, ref written);
        }

        "</borders>"u8.WriteTo(_buffer, ref span, ref written);

        _buffer.Advance(written);
    }

    public byte[] GetWrittenData()
    {
        "</styleSheet>"u8.WriteTo(_buffer);

        var preparedData = new byte[_buffer.Written];
        _buffer.FlushAll(preparedData);
        return preparedData;
    }

    public void Dispose()
        => _buffer.Dispose();

    private void AddBorder(
        Border? border,
        ReadOnlySpan<byte> empty,
        ReadOnlySpan<byte> prefix,
        ReadOnlySpan<byte> postfix,
        ref Span<byte> span,
        ref int written)
    {
        if (!border.HasValue)
        {
            empty.WriteTo(_buffer, ref span, ref written);
            return;
        }

        prefix.WriteTo(_buffer, ref span, ref written);

        " style=\""u8.WriteTo(_buffer, ref span, ref written);
        border.Value.Style.Value.WriteTo(_buffer, ref span, ref written);
        "\">"u8.WriteTo(_buffer, ref span, ref written);

        "<color rgb=\""u8.WriteTo(_buffer, ref span, ref written);
        Utf8SpanFormattableWriter.WriteValue(border.Value.Color.ToArgb(), Hex8, null, _buffer, ref span, ref written);
        "\"/>"u8.WriteTo(_buffer, ref span, ref written);

        postfix.WriteTo(_buffer, ref span, ref written);
    }
}
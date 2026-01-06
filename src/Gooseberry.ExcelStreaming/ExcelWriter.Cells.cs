using System.Drawing;
using System.Runtime.CompilerServices;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming;

public sealed partial class ExcelWriter
{
    public ExcelWriter AddCellPicture(Stream picture, PictureFormat format, Size size)
    {
        CheckWriteCell();

        _sheetDrawings.AddPicture(_sheets[^1].Id, picture, format,
            new OneCellAnchorPicturePlacementWriter(new AnchorCell(_columnCount, _rowCount - 1), size));

        _columnCount += 1;

        return this;
    }

    public ExcelWriter AddCellPicture(ReadOnlyMemory<byte> picture, PictureFormat format, Size size)
    {
        CheckWriteCell();

        _sheetDrawings.AddPicture(_sheets[^1].Id, picture, format,
            new OneCellAnchorPicturePlacementWriter(new AnchorCell(_columnCount, _rowCount - 1), size));

        _columnCount += 1;

        return this;
    }

    public ExcelWriter AddCell(
        [InterpolatedStringHandlerArgument("")]
        Utf8InterpolatedStringHandler handler,
        StyleReference? style = null)
    {
        return AddCellUtf8String(handler.GetBytes(), style);
    }

    public ExcelWriter AddCell(string? data, StyleReference? style = null)
    {
        if (data == null)
            return AddEmptyCell(style);

        AddCellImpl(data.AsSpan(), style);
        return this;
    }

    public ExcelWriter AddCell(int data, StyleReference? style = null)
    {
        AddCellNumberImpl(data, style: style);
        return this;
    }

    public ExcelWriter AddCell(int? data, StyleReference? style = null)
    {
        return data.HasValue
            ? AddCell(data.Value, style)
            : AddEmptyCell(style);
    }

    public ExcelWriter AddCell(long data, StyleReference? style = null)
    {
        AddCellNumberImpl(data, style: style);
        return this;
    }

    public ExcelWriter AddCell(long? data, StyleReference? style = null)
    {
        return data.HasValue
            ? AddCell(data.Value, style)
            : AddEmptyCell(style);
    }

    public ExcelWriter AddCell(decimal data, StyleReference? style = null)
    {
        AddCellNumberImpl(data, style: style);
        return this;
    }

    public ExcelWriter AddCell(decimal? data, StyleReference? style = null)
    {
        return data.HasValue
            ? AddCell(data.Value, style)
            : AddEmptyCell(style);
    }

    public ExcelWriter AddCell(double data, StyleReference? style = null)
    {
        AddCellNumberImpl(data, style: style);
        return this;
    }

    public ExcelWriter AddCell(double? data, StyleReference? style = null)
    {
        return data.HasValue
            ? AddCell(data.Value, style)
            : AddEmptyCell(style);
    }

    /// <summary>
    /// Format by default is StandardFormat.DayMonthYear4WithSlashes =  d/m/yyyy or mm.dd.yyyy depending on excel locale.
    /// </summary>
    public ExcelWriter AddCell(DateTime data, StyleReference? style = null)
    {
        CheckWriteCell();
        Utf8DateTimeCellWriter.Write(data, _buffer, style ?? _styles.DefaultDateStyle);
        _columnCount += 1;

        return this;
    }

    /// <summary>
    /// Format by default is StandardFormat.DayMonthYear4WithSlashes =  d/m/yyyy or mm.dd.yyyy depending on excel locale.
    /// </summary>
    public ExcelWriter AddCell(DateTime? data, StyleReference? style = null)
    {
        return data.HasValue
            ? AddCell(data.Value, style)
            : AddEmptyCell(style);
    }

    /// <summary>
    /// Format by default is StandardFormat.DayMonthYear4WithSlashes =  d/m/yyyy or mm.dd.yyyy depending on excel locale.
    /// </summary>
    public ExcelWriter AddCell(DateOnly data, StyleReference? style = null)
    {
        CheckWriteCell();
        Utf8DateTimeCellWriter.Write(data, _buffer, style ?? _styles.DefaultDateStyle);

        _columnCount += 1;
        return this;
    }

    /// <summary>
    /// Format by default is StandardFormat.DayMonthYear4WithSlashes =  d/m/yyyy or mm.dd.yyyy depending on excel locale.
    /// </summary>
    public ExcelWriter AddCell(DateOnly? data, StyleReference? style = null)
    {
        return data.HasValue
            ? AddCell(data.Value, style)
            : AddEmptyCell(style);
    }

    public ExcelWriter AddCell(char data, StyleReference? style = null)
    {
        var span = new ReadOnlySpan<char>(ref data);
        AddCellImpl(span, style);

        return this;
    }

    public ExcelWriter AddCell(ReadOnlySpan<char> data, StyleReference? style = null)
    {
        AddCellImpl(data, style);
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void AddCellImpl(ReadOnlySpan<char> data, StyleReference? style = null)
    {
        CheckWriteCell();

        StringCellWriter.Write(data, _buffer, _encoder, style);

        _columnCount += 1;
    }

    public ExcelWriter AddCellUtf8String(ReadOnlySpan<byte> data, StyleReference? style = null)
    {
        CheckWriteCell();

        StringCellWriter.WriteUtf8(data, _buffer, style);

        _columnCount += 1;
        return this;
    }

    public ExcelWriter AddCellString<T>(
        T data,
        ReadOnlySpan<char> format = default,
        IFormatProvider? formatProvider = null,
        StyleReference? style = null)
        where T : IUtf8SpanFormattable
    {
        CheckWriteCell();

        Utf8StringCellWriter.Write(data, format, formatProvider, _buffer, style);

        _columnCount += 1;
        return this;
    }

    public ExcelWriter AddCellNumber<T>(
        T data,
        ReadOnlySpan<char> format = default,
        IFormatProvider? formatProvider = null,
        StyleReference? style = null) where T : IUtf8SpanFormattable
    {
        AddCellNumberImpl(data, format, formatProvider, style);
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void AddCellNumberImpl<T>(
        T data,
        ReadOnlySpan<char> format = default,
        IFormatProvider? formatProvider = null,
        StyleReference? style = null) where T : IUtf8SpanFormattable
    {
        CheckWriteCell();

        Utf8NumberCellWriter.Write(data, format, formatProvider, _buffer, style);

        _columnCount += 1;
    }

    public ExcelWriter AddCellSharedString(string? data, StyleReference? style = null)
    {
        return data == null
            ? AddEmptyCell(style)
            : AddStringReferenceCell(_sharedStringKeeper.GetOrAdd(data), style);
    }

    public ExcelWriter AddCell(SharedStringReference sharedString, StyleReference? style = null)
    {
        if (!_sharedStringKeeper.IsValidReference(sharedString))
            throw new ArgumentException(
                "Invalid shared string reference. String not found in the table. Check sharedStringTable in ExcelWriter constructor.",
                nameof(sharedString));

        return AddStringReferenceCell(sharedString, style);
    }

    private ExcelWriter AddStringReferenceCell(SharedStringReference sharedString, StyleReference? style = null)
    {
        CheckWriteCell();
        SharedStringCellWriter.Write(sharedString, _buffer, style);

        _columnCount += 1;

        return this;
    }

    public ExcelWriter AddCell(in Hyperlink hyperlink, StyleReference? style = null)
    {
        CheckWriteCell();
        StringCellWriter.Write(hyperlink.Text, _buffer, _encoder, style ?? _styles.DefaultHyperlinkStyle);

        _columnCount += 1;
        AddHyperlink(hyperlink);

        return this;
    }

    public ExcelWriter AddEmptyCell(StyleReference? style = null)
    {
        CheckWriteCell();

        EmptyCellWriter.Write(_buffer, style);

        _columnCount += 1;

        return this;
    }

    public ExcelWriter AddEmptyCells(uint count, StyleReference? style = null)
    {
        //TODO: Optimize with r (cellIndex)
        CheckWriteCell();

        if (count == 0)
            return this;

        for (var i = 0; i < count; i++)
            EmptyCellWriter.Write(_buffer, style);

        _columnCount += count;

        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ExcelWriter MergeCells(uint colSpan = 0, uint rowSpan = 0)
    {
        if (colSpan != 0 || rowSpan != 0)
            _merges.Add(new Merge(_rowCount, _columnCount, rowSpan, colSpan));

        return this;
    }

    private void AddHyperlink(in Hyperlink hyperlink)
    {
        _hyperlinks ??= new Dictionary<string, List<CellReference>>();

        if (!_hyperlinks.TryGetValue(hyperlink.Link, out var references))
        {
            references = new List<CellReference>();
            _hyperlinks[hyperlink.Link] = references;
        }

        references.Add(new CellReference(_columnCount, _rowCount));
    }
}
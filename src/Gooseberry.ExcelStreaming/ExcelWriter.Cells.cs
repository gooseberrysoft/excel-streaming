using System.Drawing;
using System.Runtime.CompilerServices;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming;

public sealed partial class ExcelWriter
{
    public ExcelWriter AddCellPicture(Stream picture, PictureFormat format, Size size, uint rightMerge = 0, uint downMerge = 0)
    {
        CheckWriteCell();

        _sheetDrawings.AddPicture(_sheets[^1].Id, picture, format,
            new OneCellAnchorPicturePlacementWriter(new AnchorCell(_columnCount, _rowCount - 1), size));

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);

        return this;
    }

    public ExcelWriter AddCellPicture(
        ReadOnlyMemory<byte> picture,
        PictureFormat format,
        Size size,
        uint rightMerge = 0,
        uint downMerge = 0)
    {
        CheckWriteCell();

        _sheetDrawings.AddPicture(_sheets[^1].Id, picture, format,
            new OneCellAnchorPicturePlacementWriter(new AnchorCell(_columnCount, _rowCount - 1), size));

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);

        return this;
    }

#if NET8_0_OR_GREATER
    public ExcelWriter AddCell(
        [InterpolatedStringHandlerArgument("")]
        Utf8InterpolatedStringHandler handler,
        StyleReference? style = null,
        uint rightMerge = 0,
        uint downMerge = 0)
    {
        return AddCellUtf8String(handler.GetBytes(), style, rightMerge, downMerge);
    }
#endif

    public ExcelWriter AddCell(string? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        if (data == null)
            return AddEmptyCell(style, rightMerge, downMerge);

        AddCellImpl(data.AsSpan(), style, rightMerge, downMerge);
        return this;
    }

    public ExcelWriter AddCell(int data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
#if NET8_0_OR_GREATER
        AddCellNumberImpl(data, style: style, rightMerge: rightMerge, downMerge: downMerge);
#else
        CheckWriteCell();
        DataWriters.IntCellWriter.Write(data, _buffer, style);

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);
#endif
        return this;
    }

    public ExcelWriter AddCell(int? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        return data.HasValue
            ? AddCell(data.Value, style, rightMerge, downMerge)
            : AddEmptyCell(style, rightMerge, downMerge);
    }

    public ExcelWriter AddCell(long data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
#if NET8_0_OR_GREATER
        AddCellNumberImpl(data, style: style, rightMerge: rightMerge, downMerge: downMerge);
#else
        CheckWriteCell();
        DataWriters.LongCellWriter.Write(data, _buffer, style);

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);
#endif
        return this;
    }

    public ExcelWriter AddCell(long? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        return data.HasValue
            ? AddCell(data.Value, style, rightMerge, downMerge)
            : AddEmptyCell(style, rightMerge, downMerge);
    }

    public ExcelWriter AddCell(decimal data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
#if NET8_0_OR_GREATER
        AddCellNumberImpl(data, style: style, rightMerge: rightMerge, downMerge: downMerge);
#else
        CheckWriteCell();
        DataWriters.DecimalCellWriter.Write(data, _buffer, style);

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);
#endif
        return this;
    }

    public ExcelWriter AddCell(decimal? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        return data.HasValue
            ? AddCell(data.Value, style, rightMerge, downMerge)
            : AddEmptyCell(style, rightMerge, downMerge);
    }

    public ExcelWriter AddCell(double data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
#if NET8_0_OR_GREATER
        AddCellNumberImpl(data, style: style, rightMerge: rightMerge, downMerge: downMerge);
#else
        CheckWriteCell();
        DataWriters.DoubleCellWriter.Write(data, _buffer, style);

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);
#endif
        return this;
    }

    public ExcelWriter AddCell(double? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        return data.HasValue
            ? AddCell(data.Value, style, rightMerge, downMerge)
            : AddEmptyCell(style, rightMerge, downMerge);
    }

    /// <summary>
    /// Format by default is StandardFormat.DayMonthYear4WithSlashes =  d/m/yyyy or mm.dd.yyyy depending on excel locale.
    /// </summary>
    public ExcelWriter AddCell(DateTime data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        CheckWriteCell();
#if NET8_0_OR_GREATER
        Utf8DateTimeCellWriter.Write(data, _buffer, style ?? _styles.DefaultDateStyle);
#else
        DataWriters.DateTimeCellWriter.Write(data, _buffer, style ?? _styles.DefaultDateStyle);
#endif
        _columnCount += 1;
        MergeCell(rightMerge, downMerge);

        return this;
    }

    /// <summary>
    /// Format by default is StandardFormat.DayMonthYear4WithSlashes =  d/m/yyyy or mm.dd.yyyy depending on excel locale.
    /// </summary>
    public ExcelWriter AddCell(DateTime? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        return data.HasValue
            ? AddCell(data.Value, style, rightMerge, downMerge)
            : AddEmptyCell(style, rightMerge, downMerge);
    }

    /// <summary>
    /// Format by default is StandardFormat.DayMonthYear4WithSlashes =  d/m/yyyy or mm.dd.yyyy depending on excel locale.
    /// </summary>
    public ExcelWriter AddCell(DateOnly data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
#if NET8_0_OR_GREATER
        CheckWriteCell();
        Utf8DateTimeCellWriter.Write(data, _buffer, style ?? _styles.DefaultDateStyle);

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);
#else
        AddCell(data.ToDateTime(default), style, rightMerge, downMerge);
#endif
        return this;
    }

    /// <summary>
    /// Format by default is StandardFormat.DayMonthYear4WithSlashes =  d/m/yyyy or mm.dd.yyyy depending on excel locale.
    /// </summary>
    public ExcelWriter AddCell(DateOnly? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        return data.HasValue
            ? AddCell(data.Value, style, rightMerge, downMerge)
            : AddEmptyCell(style, rightMerge, downMerge);
    }

    public ExcelWriter AddCell(char data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
#if NET8_0_OR_GREATER
        var span = new ReadOnlySpan<char>(ref data);
#else
        ReadOnlySpan<char> span = stackalloc char[] { data };
#endif
        AddCellImpl(span, style, rightMerge, downMerge);

        return this;
    }

    public ExcelWriter AddCell(ReadOnlySpan<char> data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        AddCellImpl(data, style, rightMerge, downMerge);
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void AddCellImpl(ReadOnlySpan<char> data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        CheckWriteCell();

        StringCellWriter.Write(data, _buffer, _encoder, style);

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);
    }

    public ExcelWriter AddCellUtf8String(ReadOnlySpan<byte> data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        CheckWriteCell();

        StringCellWriter.WriteUtf8(data, _buffer, style);

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);
        return this;
    }

#if NET8_0_OR_GREATER
    public ExcelWriter AddCellString<T>(
        T data,
        ReadOnlySpan<char> format = default,
        IFormatProvider? formatProvider = default,
        StyleReference? style = null,
        uint rightMerge = 0,
        uint downMerge = 0)
        where T : IUtf8SpanFormattable
    {
        CheckWriteCell();

        Utf8StringCellWriter.Write(data, format, formatProvider, _buffer, style);

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);
        return this;
    }

    public ExcelWriter AddCellNumber<T>(
        T data,
        ReadOnlySpan<char> format = default,
        IFormatProvider? formatProvider = default,
        StyleReference? style = null,
        uint rightMerge = 0,
        uint downMerge = 0) where T : IUtf8SpanFormattable
    {
        AddCellNumberImpl(data, format, formatProvider, style, rightMerge, downMerge);
        return this;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void AddCellNumberImpl<T>(
        T data,
        ReadOnlySpan<char> format = default,
        IFormatProvider? formatProvider = default,
        StyleReference? style = null,
        uint rightMerge = 0,
        uint downMerge = 0) where T : IUtf8SpanFormattable
    {
        CheckWriteCell();

        Utf8NumberCellWriter.Write(data, format, formatProvider, _buffer, style);

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);
    }
#endif

    public ExcelWriter AddCellSharedString(string? data, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        return data == null
            ? AddEmptyCell(style, rightMerge, downMerge)
            : AddStringReferenceCell(_sharedStringKeeper.GetOrAdd(data), style, rightMerge, downMerge);
    }

    public ExcelWriter AddCell(SharedStringReference sharedString, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        if (!_sharedStringKeeper.IsValidReference(sharedString))
            throw new ArgumentException(
                "Invalid shared string reference. String not found in the table. Check sharedStringTable in ExcelWriter constructor.",
                nameof(sharedString));

        return AddStringReferenceCell(sharedString, style, rightMerge, downMerge);
    }

    private ExcelWriter AddStringReferenceCell(
        SharedStringReference sharedString,
        StyleReference? style = null,
        uint rightMerge = 0,
        uint downMerge = 0)
    {
        CheckWriteCell();
        SharedStringCellWriter.Write(sharedString, _buffer, style);

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);

        return this;
    }

    public ExcelWriter AddCell(in Hyperlink hyperlink, StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        CheckWriteCell();
        StringCellWriter.Write(hyperlink.Text, _buffer, _encoder, style ?? _styles.DefaultHyperlinkStyle);

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);
        AddHyperlink(hyperlink);

        return this;
    }

    public ExcelWriter AddEmptyCell(StyleReference? style = null, uint rightMerge = 0, uint downMerge = 0)
    {
        CheckWriteCell();

        EmptyCellWriter.Write(_buffer, style);

        _columnCount += 1;
        MergeCell(rightMerge, downMerge);

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
    private void MergeCell(uint rightMerge = 0, uint downMerge = 0)
    {
        if (rightMerge != 0 || downMerge != 0)
            AddMerge(rightMerge, downMerge);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private void AddMerge(uint rightMerge, uint downMerge)
        => _merges.Add(new Merge(_rowCount, _columnCount, downMerge, rightMerge));

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
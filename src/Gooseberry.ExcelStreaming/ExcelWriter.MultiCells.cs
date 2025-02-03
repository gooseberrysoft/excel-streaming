using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming;

public sealed partial class ExcelWriter
{
    public ExcelWriter AddCells(IEnumerable<string?> cells, StyleReference? style = null)
    {
        CheckWriteCell();

        foreach (var cell in cells)
        {
            if (cell == null)
                EmptyCellWriter.Write(_buffer, style);
            else
                StringCellWriter.Write(cell.AsSpan(), _buffer, _encoder, style);

            _columnCount += 1;
        }

        return this;
    }

    public ExcelWriter AddCells(ReadOnlySpan<string?> cells, StyleReference? style = null)
    {
        CheckWriteCell();

        foreach (var cell in cells)
        {
            if (cell == null)
                EmptyCellWriter.Write(_buffer, style);
            else
                StringCellWriter.Write(cell.AsSpan(), _buffer, _encoder, style);
        }

        _columnCount += (uint)cells.Length;

        return this;
    }

    public ExcelWriter AddCells(ReadOnlySpan<(string?, StyleReference?)> cells)
    {
        CheckWriteCell();

        foreach (ref readonly var cell in cells)
        {
            if (cell.Item1 == null)
                EmptyCellWriter.Write(_buffer, cell.Item2);
            else
                StringCellWriter.Write(cell.Item1.AsSpan(), _buffer, _encoder, cell.Item2);
        }

        _columnCount += (uint)cells.Length;

        return this;
    }

    public ExcelWriter AddCells(IEnumerable<DateTime> cells, StyleReference? style = null)
    {
        CheckWriteCell();

        foreach (var cell in cells)
        {
#if NET8_0_OR_GREATER
            Utf8DateTimeCellWriter.Write(cell, _buffer, style ?? _styles.DefaultDateStyle);
#else
            DataWriters.DateTimeCellWriter.Write(cell, _buffer, style ?? _styles.DefaultDateStyle);
#endif
            _columnCount += 1;
        }

        return this;
    }

    public ExcelWriter AddCells(ReadOnlySpan<DateTime> cells, StyleReference? style = null)
    {
        CheckWriteCell();

        foreach (var cell in cells)
        {
#if NET8_0_OR_GREATER
            Utf8DateTimeCellWriter.Write(cell, _buffer, style ?? _styles.DefaultDateStyle);
#else
            DataWriters.DateTimeCellWriter.Write(cell, _buffer, style ?? _styles.DefaultDateStyle);
#endif
        }

        _columnCount += (uint)cells.Length;

        return this;
    }

    public ExcelWriter AddCells(ReadOnlySpan<(DateTime, StyleReference?)> cells)
    {
        CheckWriteCell();

        foreach (ref readonly var cell in cells)
        {
#if NET8_0_OR_GREATER
            Utf8DateTimeCellWriter.Write(cell.Item1, _buffer, cell.Item2 ?? _styles.DefaultDateStyle);
#else
            DataWriters.DateTimeCellWriter.Write(cell.Item1, _buffer, cell.Item2 ?? _styles.DefaultDateStyle);
#endif
        }

        _columnCount += (uint)cells.Length;

        return this;
    }

    public ExcelWriter AddCells(IEnumerable<DateOnly> cells, StyleReference? style = null)
    {
        CheckWriteCell();

        foreach (var cell in cells)
        {
#if NET8_0_OR_GREATER
            Utf8DateTimeCellWriter.Write(cell, _buffer, style ?? _styles.DefaultDateStyle);
#else
            DataWriters.DateTimeCellWriter.Write(cell.ToDateTime(default), _buffer, style ?? _styles.DefaultDateStyle);
#endif
            _columnCount += 1;
        }

        return this;
    }

    public ExcelWriter AddCells(ReadOnlySpan<DateOnly> cells, StyleReference? style = null)
    {
        CheckWriteCell();

        foreach (var cell in cells)
        {
#if NET8_0_OR_GREATER
            Utf8DateTimeCellWriter.Write(cell, _buffer, style ?? _styles.DefaultDateStyle);
#else
            DataWriters.DateTimeCellWriter.Write(cell.ToDateTime(default), _buffer, style ?? _styles.DefaultDateStyle);
#endif
        }

        _columnCount += (uint)cells.Length;

        return this;
    }

    public ExcelWriter AddCells(ReadOnlySpan<(DateOnly, StyleReference?)> cells)
    {
        CheckWriteCell();

        foreach (ref readonly var cell in cells)
        {
#if NET8_0_OR_GREATER
            Utf8DateTimeCellWriter.Write(cell.Item1, _buffer, cell.Item2 ?? _styles.DefaultDateStyle);
#else
            DataWriters.DateTimeCellWriter.Write(cell.Item1.ToDateTime(default), _buffer, cell.Item2 ?? _styles.DefaultDateStyle);
#endif
        }

        _columnCount += (uint)cells.Length;

        return this;
    }

#if NET8_0_OR_GREATER
    public ExcelWriter AddCellsString<T>(
        ReadOnlySpan<T> cells,
        ReadOnlySpan<char> format = default,
        IFormatProvider? formatProvider = default,
        StyleReference? style = null)
        where T : IUtf8SpanFormattable
    {
        CheckWriteCell();

        foreach (ref readonly var cell in cells)
            Utf8StringCellWriter.Write(cell, format, formatProvider, _buffer, style);

        _columnCount += (uint)cells.Length;

        return this;
    }

    public ExcelWriter AddCellsString<T>(
        ReadOnlySpan<(T, StyleReference?)> cells,
        ReadOnlySpan<char> format = default,
        IFormatProvider? formatProvider = default)
        where T : IUtf8SpanFormattable
    {
        CheckWriteCell();

        foreach (ref readonly var cell in cells)
            Utf8StringCellWriter.Write(cell.Item1, format, formatProvider, _buffer, cell.Item2);

        _columnCount += (uint)cells.Length;

        return this;
    }

    public ExcelWriter AddCellsString<T>(
        IEnumerable<T> cells,
        ReadOnlySpan<char> format = default,
        IFormatProvider? formatProvider = default,
        StyleReference? style = null)
        where T : IUtf8SpanFormattable
    {
        CheckWriteCell();

        foreach (var cell in cells)
        {
            Utf8StringCellWriter.Write(cell, format, formatProvider, _buffer, style);
            _columnCount += 1;
        }

        return this;
    }

    public ExcelWriter AddCellsNumber<T>(
        ReadOnlySpan<T> cells,
        ReadOnlySpan<char> format = default,
        IFormatProvider? formatProvider = default,
        StyleReference? style = null) where T : IUtf8SpanFormattable
    {
        CheckWriteCell();

        foreach (ref readonly var cell in cells)
            Utf8NumberCellWriter.Write(cell, format, formatProvider, _buffer, style);

        _columnCount += (uint)cells.Length;
        return this;
    }

    public ExcelWriter AddCellsNumber<T>(
        ReadOnlySpan<(T, StyleReference?)> cells,
        ReadOnlySpan<char> format = default,
        IFormatProvider? formatProvider = default) where T : IUtf8SpanFormattable
    {
        CheckWriteCell();

        foreach (ref readonly var cell in cells)
            Utf8NumberCellWriter.Write(cell.Item1, format, formatProvider, _buffer, cell.Item2);

        _columnCount += (uint)cells.Length;
        return this;
    }

    public ExcelWriter AddCellsNumber<T>(
        IEnumerable<T> cells,
        ReadOnlySpan<char> format = default,
        IFormatProvider? formatProvider = default,
        StyleReference? style = null) where T : IUtf8SpanFormattable
    {
        CheckWriteCell();

        foreach (var cell in cells)
        {
            Utf8NumberCellWriter.Write(cell, format, formatProvider, _buffer, style);
            _columnCount += 1;
        }

        return this;
    }

#endif

    public ExcelWriter AddCellsUtf8String(
        ReadOnlySpan<byte> cell1,
        ReadOnlySpan<byte> cell2,
        StyleReference? style = null)
    {
        CheckWriteCell();

        StringCellWriter.WriteUtf8(cell1, _buffer, style);
        StringCellWriter.WriteUtf8(cell2, _buffer, style);

        _columnCount += 2;
        return this;
    }

    public ExcelWriter AddCellsUtf8String(
        ReadOnlySpan<byte> cell1,
        ReadOnlySpan<byte> cell2,
        ReadOnlySpan<byte> cell3,
        StyleReference? style = null)
    {
        CheckWriteCell();

        StringCellWriter.WriteUtf8(cell1, _buffer, style);
        StringCellWriter.WriteUtf8(cell2, _buffer, style);
        StringCellWriter.WriteUtf8(cell3, _buffer, style);

        _columnCount += 3;
        return this;
    }

    public ExcelWriter AddCellsUtf8String(
        ReadOnlySpan<byte> cell1,
        ReadOnlySpan<byte> cell2,
        ReadOnlySpan<byte> cell3,
        ReadOnlySpan<byte> cell4,
        StyleReference? style = null)
    {
        CheckWriteCell();

        StringCellWriter.WriteUtf8(cell1, _buffer, style);
        StringCellWriter.WriteUtf8(cell2, _buffer, style);
        StringCellWriter.WriteUtf8(cell3, _buffer, style);
        StringCellWriter.WriteUtf8(cell4, _buffer, style);

        _columnCount += 4;
        return this;
    }

    public ExcelWriter AddCellsUtf8String(
        ReadOnlySpan<byte> cell1,
        ReadOnlySpan<byte> cell2,
        ReadOnlySpan<byte> cell3,
        ReadOnlySpan<byte> cell4,
        ReadOnlySpan<byte> cell5,
        StyleReference? style = null)
    {
        CheckWriteCell();

        StringCellWriter.WriteUtf8(cell1, _buffer, style);
        StringCellWriter.WriteUtf8(cell2, _buffer, style);
        StringCellWriter.WriteUtf8(cell3, _buffer, style);
        StringCellWriter.WriteUtf8(cell4, _buffer, style);
        StringCellWriter.WriteUtf8(cell5, _buffer, style);

        _columnCount += 5;
        return this;
    }

    public ExcelWriter AddCellsUtf8String(
        ReadOnlySpan<byte> cell1,
        ReadOnlySpan<byte> cell2,
        ReadOnlySpan<byte> cell3,
        ReadOnlySpan<byte> cell4,
        ReadOnlySpan<byte> cell5,
        ReadOnlySpan<byte> cell6,
        StyleReference? style = null)
    {
        CheckWriteCell();

        StringCellWriter.WriteUtf8(cell1, _buffer, style);
        StringCellWriter.WriteUtf8(cell2, _buffer, style);
        StringCellWriter.WriteUtf8(cell3, _buffer, style);
        StringCellWriter.WriteUtf8(cell4, _buffer, style);
        StringCellWriter.WriteUtf8(cell5, _buffer, style);
        StringCellWriter.WriteUtf8(cell6, _buffer, style);
        
        _columnCount += 6;
        return this;
    }

    public ExcelWriter AddCellsUtf8String(
        ReadOnlySpan<byte> cell1,
        ReadOnlySpan<byte> cell2,
        ReadOnlySpan<byte> cell3,
        ReadOnlySpan<byte> cell4,
        ReadOnlySpan<byte> cell5,
        ReadOnlySpan<byte> cell6,
        ReadOnlySpan<byte> cell7,
        StyleReference? style = null)
    {
        CheckWriteCell();

        StringCellWriter.WriteUtf8(cell1, _buffer, style);
        StringCellWriter.WriteUtf8(cell2, _buffer, style);
        StringCellWriter.WriteUtf8(cell3, _buffer, style);
        StringCellWriter.WriteUtf8(cell4, _buffer, style);
        StringCellWriter.WriteUtf8(cell5, _buffer, style);
        StringCellWriter.WriteUtf8(cell6, _buffer, style);
        StringCellWriter.WriteUtf8(cell7, _buffer, style);

        _columnCount += 7;
        return this;
    }
}
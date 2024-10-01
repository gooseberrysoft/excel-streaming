using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

using RowCellConstants = Constants.Worksheet.SheetData.Row.Cell;

internal sealed class NumberCellWriter<T, TFormatter>
    where TFormatter : INumberFormatter<T>, new()
{
    private readonly NumberWriter<T, TFormatter> _valueWriter = new();

    private readonly byte[] _stylelessPrefix;
    private readonly byte[] _stylePrefix;
    private readonly byte[] _stylePostfix;

    public NumberCellWriter(ReadOnlySpan<byte> dataType)
    {
        _stylelessPrefix = RowCellConstants.Prefix.Combine(dataType, RowCellConstants.Middle);

        _stylePrefix = RowCellConstants.Prefix.Combine(dataType, RowCellConstants.Style.Prefix);

        _stylePostfix = RowCellConstants.Style.Postfix.Combine(RowCellConstants.Middle);
    }

    public void Write(in T value, BuffersChain buffer, StyleReference? style = null)
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            _stylePrefix.WriteTo(buffer, ref span, ref written);
            style.Value.Value.WriteTo(buffer, ref span, ref written);
            _stylePostfix.WriteTo(buffer, ref span, ref written);
            _valueWriter.WriteValue(value, buffer, ref span, ref written);
            RowCellConstants.Postfix.WriteTo(buffer, ref span, ref written);

            buffer.Advance(written);
            return;
        }

        _stylelessPrefix.WriteTo(buffer, ref span, ref written);
        _valueWriter.WriteValue(value, buffer, ref span, ref written);
        RowCellConstants.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
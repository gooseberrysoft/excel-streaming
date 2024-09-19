using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class NumberCellWriter<T, TFormatter>
    where TFormatter : INumberFormatter<T>, new()
{
    private readonly NumberWriter<int, IntFormatter> _styleWriter = new();
    private readonly NumberWriter<T, TFormatter> _valueWriter = new();

    private readonly byte[] _stylelessPrefix;
    private readonly byte[] _stylePrefix;
    private readonly byte[] _stylePostfix;

    public NumberCellWriter(ReadOnlySpan<byte> dataType)
    {
        _stylelessPrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
            .Combine(dataType, Constants.Worksheet.SheetData.Row.Cell.Middle);

        _stylePrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
            .Combine(dataType, Constants.Worksheet.SheetData.Row.Cell.Style.Prefix);

        _stylePostfix = Constants.Worksheet.SheetData.Row.Cell.Style.Postfix
            .Combine(Constants.Worksheet.SheetData.Row.Cell.Middle);
    }

    public void Write(in T value, BuffersChain buffer, StyleReference? style = null)
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            _stylePrefix.WriteTo(buffer, ref span, ref written);
            _styleWriter.WriteValue(style.Value.Value, buffer, ref span, ref written);
            _stylePostfix.WriteTo(buffer, ref span, ref written);
            _valueWriter.WriteValue(value, buffer, ref span, ref written);
            Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);

            buffer.Advance(written);
            return;
        }

        _stylelessPrefix.WriteTo(buffer, ref span, ref written);
        _valueWriter.WriteValue(value, buffer, ref span, ref written);
        Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}

internal static class Utf8NumberCellWriter
{
    private static readonly byte[] _stylelessPrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
        .Combine(Constants.Worksheet.SheetData.Row.Cell.NumberDataType, Constants.Worksheet.SheetData.Row.Cell.Middle);

    private static readonly byte[] _stylePrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
        .Combine(Constants.Worksheet.SheetData.Row.Cell.NumberDataType, Constants.Worksheet.SheetData.Row.Cell.Style.Prefix);

    private static readonly byte[] _stylePostfix = Constants.Worksheet.SheetData.Row.Cell.Style.Postfix
        .Combine(Constants.Worksheet.SheetData.Row.Cell.Middle);

    public static void Write<T>(in T value, BuffersChain buffer, StyleReference? style = null)
        where T : IUtf8SpanFormattable
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            _stylePrefix.WriteTo(buffer, ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(style.Value.Value, buffer, ref span, ref written);
            _stylePostfix.WriteTo(buffer, ref span, ref written);
            _valueWriter.WriteValue(value, buffer, ref span, ref written);
            Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);

            buffer.Advance(written);
            return;
        }

        _stylelessPrefix.WriteTo(buffer, ref span, ref written);
        _valueWriter.WriteValue(value, buffer, ref span, ref written);
        Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
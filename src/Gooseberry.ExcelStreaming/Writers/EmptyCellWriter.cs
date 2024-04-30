using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class EmptyCellWriter
{
    private readonly byte[] _stateless;
    private readonly byte[] _stylePrefix;
    private readonly byte[] _stylePostfix;

    private readonly NumberWriter<int, IntFormatter> _styleWriter = new();

    public EmptyCellWriter()
    {
        _stateless = Constants.Worksheet.SheetData.Row.Cell.Prefix
            .Combine(Constants.Worksheet.SheetData.Row.Cell.StringDataType,
            Constants.Worksheet.SheetData.Row.Cell.Middle,
            Constants.Worksheet.SheetData.Row.Cell.Postfix);

        _stylePrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
            .Combine(Constants.Worksheet.SheetData.Row.Cell.StringDataType,
                Constants.Worksheet.SheetData.Row.Cell.Style.Prefix);

        _stylePostfix = Constants.Worksheet.SheetData.Row.Cell.Style.Postfix
            .Combine(Constants.Worksheet.SheetData.Row.Cell.Middle);
    }

    public void Write(BuffersChain buffer, StyleReference? style = null)
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            _stylePrefix.WriteTo(buffer, ref span, ref written);
            _styleWriter.WriteValue(style.Value.Value, buffer, ref span, ref written);
            _stylePostfix.WriteTo(buffer, ref span, ref written);
            Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);

            buffer.Advance(written);
            return;
        }

        _stateless.WriteTo(buffer, ref span, ref written);
        buffer.Advance(written);
    }
}
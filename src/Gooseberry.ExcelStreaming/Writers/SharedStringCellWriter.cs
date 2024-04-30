using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class SharedStringCellWriter
{
    private readonly NumberWriter<int, IntFormatter> _intWriter = new();

    private readonly byte[] _stylelessPrefix;
    private readonly byte[] _stylePrefix;
    private readonly byte[] _stylePostfix;

    public SharedStringCellWriter()
    {
        _stylelessPrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
            .Combine(Constants.Worksheet.SheetData.Row.Cell.SharedStringDataType,
            Constants.Worksheet.SheetData.Row.Cell.Middle);

        _stylePrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
            .Combine(Constants.Worksheet.SheetData.Row.Cell.SharedStringDataType,
            Constants.Worksheet.SheetData.Row.Cell.Style.Prefix);

        _stylePostfix = Constants.Worksheet.SheetData.Row.Cell.Style.Postfix
            .Combine(Constants.Worksheet.SheetData.Row.Cell.Middle);
    }

    public void Write(SharedStringReference sharedString, BuffersChain buffer, StyleReference? style = null)
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            _stylePrefix.WriteTo(buffer, ref span, ref written);
            _intWriter.WriteValue(style.Value.Value, buffer, ref span, ref written);
            _stylePostfix.WriteTo(buffer, ref span, ref written);
            _intWriter.WriteValue(sharedString.Value, buffer, ref span, ref written);
            Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);

            buffer.Advance(written);
            return;
        }

        _stylelessPrefix.WriteTo(buffer, ref span, ref written);
        _intWriter.WriteValue(sharedString.Value, buffer, ref span, ref written);
        Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
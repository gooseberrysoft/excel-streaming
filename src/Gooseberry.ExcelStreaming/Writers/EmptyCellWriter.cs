using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Styles;
using Gooseberry.ExcelStreaming.Writers.Formatters;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class EmptyCellWriter
{
    private static readonly byte[] _stateless;
    private static readonly byte[] _stylePrefix;
    private static readonly byte[] _stylePostfix;
    
    static EmptyCellWriter()
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

    public static void Write(BuffersChain buffer, StyleReference? style = null)
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            _stylePrefix.WriteTo(buffer, ref span, ref written);
            NumberWriter<int, IntFormatter>.WriteValue(style.Value.Value, buffer, ref span, ref written);
            _stylePostfix.WriteTo(buffer, ref span, ref written);
            Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);

            buffer.Advance(written);
            return;
        }

        _stateless.WriteTo(buffer, ref span, ref written);
        buffer.Advance(written);
    }
}
#if NET8_0_OR_GREATER
using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class Utf8StringCellWriter
{
    private static readonly byte[] StylelessPrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
        .Combine(Constants.Worksheet.SheetData.Row.Cell.StringDataType, Constants.Worksheet.SheetData.Row.Cell.Middle);

    private static readonly byte[] StylePrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
        .Combine(Constants.Worksheet.SheetData.Row.Cell.StringDataType, Constants.Worksheet.SheetData.Row.Cell.Style.Prefix);

    private static readonly byte[] StylePostfix = Constants.Worksheet.SheetData.Row.Cell.Style.Postfix
        .Combine(Constants.Worksheet.SheetData.Row.Cell.Middle);

    public static void Write<T>(
        T value,
        ReadOnlySpan<char> format,
        IFormatProvider? provider,
        BuffersChain buffer,
        StyleReference? style = null)
        where T : IUtf8SpanFormattable
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            StylePrefix.WriteTo(buffer, ref span, ref written);
            style.Value.Value.WriteTo(buffer, ref span, ref written);
            StylePostfix.WriteTo(buffer, ref span, ref written);
            StringWriter.WriteEscapedUtf8To(value, format, provider, buffer, ref span, ref written);
            Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);

            buffer.Advance(written);
            return;
        }

        StylelessPrefix.WriteTo(buffer, ref span, ref written);
        StringWriter.WriteEscapedUtf8To(value, format, provider, buffer, ref span, ref written);
        Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
#endif
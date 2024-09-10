using Gooseberry.ExcelStreaming.Extensions;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class RowWriter
{
    private readonly NumberWriter<decimal, DecimalFormatter> _rowNumberWriter = new();

    private static readonly byte[] RowCloseAndStart =
        Constants.Worksheet.SheetData.Row.Postfix
            .Combine(Constants.Worksheet.SheetData.Row.Open.Prefix,
                Constants.Worksheet.SheetData.Row.Open.Postfix);

    private static readonly byte[] RowStart =
        Constants.Worksheet.SheetData.Row.Open.Prefix
            .Combine(Constants.Worksheet.SheetData.Row.Open.Postfix);

    private static readonly byte[] RowHeightPrefix =
        Constants.Worksheet.SheetData.Row.Open.Prefix
            .Combine(Constants.Worksheet.SheetData.Row.Open.Height.Prefix);

    private static readonly byte[] RowHeightPostfix =
        Constants.Worksheet.SheetData.Row.Open.Height.Postfix
            .Combine(Constants.Worksheet.SheetData.Row.Open.Postfix);

    public void WriteStartRow(
        BuffersChain buffer,
        bool rowStarted,
        decimal? height = null,
        bool isCollapsed = false,
        uint? outlineLevel = null,
        bool isHidden = false)
    {
        var span = buffer.GetSpan();
        var written = 0;

        var rowHasAttributes = HasAttributes(height, isCollapsed, outlineLevel, isHidden);

        if (rowStarted && !rowHasAttributes)
        {
            RowCloseAndStart.WriteTo(buffer, ref span, ref written);
            buffer.Advance(written);

            return;
        }

        if (rowStarted)
        {
            Constants.Worksheet.SheetData.Row.Postfix.WriteTo(buffer, ref span, ref written);
            Constants.Worksheet.SheetData.Row.Open.Prefix.WriteTo(buffer, ref span, ref written);
        }
        else
            Constants.Worksheet.SheetData.Row.Open.Prefix.WriteTo(buffer, ref span, ref written);

        if (!rowHasAttributes)
        {
            RowStart.WriteTo(buffer, ref span, ref written);
            buffer.Advance(written);

            return;
        }

        if (height.HasValue)
        {
            Constants.Worksheet.SheetData.Row.Open.Height.Prefix.WriteTo(buffer, ref span, ref written);
            _rowNumberWriter.WriteValue(height.Value, buffer, ref span, ref written);
            Constants.Worksheet.SheetData.Row.Open.Height.Postfix.WriteTo(buffer, ref span, ref written);
        }

        if (isHidden)
            Constants.Worksheet.SheetData.Row.Open.Hidden.WriteTo(buffer, ref span, ref written);

        if (outlineLevel.HasValue)
        {
            Constants.Worksheet.SheetData.Row.Open.OutlineLevel.Prefix.WriteTo(buffer, ref span, ref written);
            _rowNumberWriter.WriteValue(outlineLevel.Value, buffer, ref span, ref written);
            Constants.Worksheet.SheetData.Row.Open.OutlineLevel.Postfix.WriteTo(buffer, ref span, ref written);
        }

        if (isCollapsed)
            Constants.Worksheet.SheetData.Row.Open.Collapsed.WriteTo(buffer, ref span, ref written);

        Constants.Worksheet.SheetData.Row.Open.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }

    private static bool HasAttributes(
        decimal? height = null,
        bool isCollapsed = false,
        uint? outlineLevel = null,
        bool isHidden = false)
    {
        return height.HasValue || isCollapsed || outlineLevel.HasValue || isHidden;
    }

    public void WriteEndRow(BuffersChain buffer)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.Worksheet.SheetData.Row.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
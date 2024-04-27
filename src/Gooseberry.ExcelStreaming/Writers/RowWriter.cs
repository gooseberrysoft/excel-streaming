using Gooseberry.ExcelStreaming.Extensions;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class RowWriter
{
    private readonly NumberWriter<decimal, DecimalFormatter> _rowHeightWriter = new();

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

    public void WriteStartRow(BuffersChain buffer, bool rowStarted, decimal? height = null)
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (rowStarted && !height.HasValue)
        {
            RowCloseAndStart.WriteTo(buffer, ref span, ref written);
            buffer.Advance(written);
            return;
        }

        if (rowStarted)
            Constants.Worksheet.SheetData.Row.Postfix.WriteTo(buffer, ref span, ref written);

        if (height.HasValue)
        {
            RowHeightPrefix.WriteTo(buffer, ref span, ref written);
            _rowHeightWriter.WriteValue(height.Value, buffer, ref span, ref written);
            RowHeightPostfix.WriteTo(buffer, ref span, ref written);
        }
        else
            RowStart.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }

    public void WriteEndRow(BuffersChain buffer)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.Worksheet.SheetData.Row.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
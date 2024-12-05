using Gooseberry.ExcelStreaming.Writers.Formatters;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class DataWriters
{
    public static readonly NumberCellWriter<int, IntFormatter> IntCellWriter =
        new(Constants.Worksheet.SheetData.Row.Cell.NumberDataType);

    public static readonly NumberCellWriter<long, LongFormatter> LongCellWriter =
        new(Constants.Worksheet.SheetData.Row.Cell.NumberDataType);

    public static readonly NumberCellWriter<decimal, DecimalFormatter> DecimalCellWriter =
        new(Constants.Worksheet.SheetData.Row.Cell.NumberDataType);

    public static readonly NumberCellWriter<double, DoubleFormatter> DoubleCellWriter =
        new(Constants.Worksheet.SheetData.Row.Cell.NumberDataType);

    public static readonly NumberCellWriter<DateTime, DateTimeFormatter> DateTimeCellWriter =
        new(Constants.Worksheet.SheetData.Row.Cell.DateTimeDataType);

    public static readonly NumberCellWriter<DateOnly, DateOnlyFormatter> DateOnlyCellWriter =
        new(Constants.Worksheet.SheetData.Row.Cell.NumberDataType);
}
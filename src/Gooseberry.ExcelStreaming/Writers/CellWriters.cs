using System;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class CellWriters
{
    public static readonly RowWriter RowWriter = new();
    
    public static readonly NumberCellWriter<int, IntFormatter> IntCellWriter =
        new(Constants.Worksheet.SheetData.Row.Cell.NumberDataType);
    
    public static readonly NumberCellWriter<long, LongFormatter> LongCellWriter =
        new(Constants.Worksheet.SheetData.Row.Cell.NumberDataType);
    
    public static readonly NumberCellWriter<decimal, DecimalFormatter> DecimalCellWriter =
        new(Constants.Worksheet.SheetData.Row.Cell.NumberDataType);
    
    public static readonly NumberCellWriter<DateTime, DateTimeFormatter> DateTimeCellWriter =
        new(Constants.Worksheet.SheetData.Row.Cell.DateTimeDataType);

    public static readonly StringCellWriter StringCellWriter = new();

    public static readonly EmptyCellWriter EmptyCellWriter = new();
}
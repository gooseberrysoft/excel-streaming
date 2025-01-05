using System.Text;
using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

using RowCellConstants = Constants.Worksheet.SheetData.Row.Cell;

internal static class StringCellWriter
{
    // https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3?ui=en-us&rs=en-us&ad=us#ID0EBABAAA=Excel_2016-2013
    private const int MaxCharacters = 32_767;
    internal const int MaxBytes = MaxCharacters * 3;

    private static readonly byte[] _stylelessPrefix;
    private static readonly byte[] _stylePrefix;
    private static readonly byte[] _stylePostfix;


    static StringCellWriter()
    {
        _stylelessPrefix = RowCellConstants.Prefix
            .Combine(RowCellConstants.StringDataType,
                RowCellConstants.Middle);

        _stylePrefix = RowCellConstants.Prefix
            .Combine(RowCellConstants.StringDataType,
                RowCellConstants.Style.Prefix);

        _stylePostfix = RowCellConstants.Style.Postfix
            .Combine(RowCellConstants.Middle);
    }

    public static void Write(ReadOnlySpan<char> value, BuffersChain buffer, Encoder encoder, StyleReference? style = null)
    {
        if (value.Length > MaxCharacters)
            ThrowCharsLimitExceeded();

        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            _stylePrefix.WriteTo(buffer, ref span, ref written);
            style.Value.Value.WriteTo(buffer, ref span, ref written);
            _stylePostfix.WriteTo(buffer, ref span, ref written);
            value.WriteEscapedTo(buffer, encoder, ref span, ref written);
            RowCellConstants.Postfix.WriteTo(buffer, ref span, ref written);

            buffer.Advance(written);
            return;
        }

        _stylelessPrefix.WriteTo(buffer, ref span, ref written);
        value.WriteEscapedTo(buffer, encoder, ref span, ref written);
        RowCellConstants.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }

    public static void WriteUtf8(ReadOnlySpan<byte> value, BuffersChain buffer, StyleReference? style = null)
    {
        if (value.Length > MaxBytes)
            ThrowCharsLimitExceeded();

        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            _stylePrefix.WriteTo(buffer, ref span, ref written);
            style.Value.Value.WriteTo(buffer, ref span, ref written);
            _stylePostfix.WriteTo(buffer, ref span, ref written);
            value.WriteEscapedUtf8To(buffer, ref span, ref written);
            RowCellConstants.Postfix.WriteTo(buffer, ref span, ref written);

            buffer.Advance(written);
            return;
        }

        _stylelessPrefix.WriteTo(buffer, ref span, ref written);
        value.WriteEscapedUtf8To(buffer, ref span, ref written);
        RowCellConstants.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }

    public static void ThrowCharsLimitExceeded()
        => throw new ArgumentException($"Cell value exceed Excel {MaxCharacters} chars limit.");
}
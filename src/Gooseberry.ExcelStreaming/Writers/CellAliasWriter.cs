using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class CellAliasWriter
{
    private static readonly byte[] Colon = [(byte)':'];
    private const int ColumnNameMaxLength = 3; //XFD last column name
    private const int RowMaxLength = 7; //1000000 
    private const int MaxLength = ColumnNameMaxLength + RowMaxLength;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this in Merge merge,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
    {
        WriteTo(merge.RightBottom, buffer, ref destination, ref written);
        Colon.WriteTo(buffer, ref destination, ref written);
        WriteTo(merge.TopLeft, buffer, ref destination, ref written);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this CellReference cellReference,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
    {
        if (destination.Length < MaxLength)
        {
            buffer.Advance(written);
            destination = buffer.GetSpan(MaxLength);
            written = 0;
        }

        WriteColumnName(cellReference.Column, ref destination, ref written);
        cellReference.Row.WriteTo(buffer, ref destination, ref written);
    }

    private static void WriteColumnName(uint columnNumber, ref Span<byte> destination, ref int written)
    {
        if (columnNumber <= 26)
        {
            destination[0] = Convert.ToByte('A' + (columnNumber - 1));
            destination = destination.Slice(1);
            written += 1;

            return;
        }

        WriteMultiSymbolColumnName(columnNumber, ref destination, ref written);
    }

    private static void WriteMultiSymbolColumnName(uint columnNumber, ref Span<byte> destination, ref int written)
    {
        Span<byte> columnName = stackalloc byte[ColumnNameMaxLength];
        int index = 0;

        while (columnNumber > 0)
        {
            var modulo = (columnNumber - 1) % 26;
            columnName[index] = Convert.ToByte('A' + modulo);
            columnNumber = (columnNumber - modulo) / 26;
            index++;
        }

        columnName = columnName[..index];
        columnName.Reverse();

        columnName.CopyTo(destination);

        destination = destination.Slice(columnName.Length);
        written += columnName.Length;
    }
}
using System.Runtime.InteropServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

[StructLayout(LayoutKind.Auto)]
public readonly struct CellReference
{
    public CellReference(uint column, uint row)
    {
        CheckValidCell(column, row);
        Row = row;
        Column = column;
    }

    internal readonly uint Column;
    internal readonly uint Row;

    public static implicit operator CellReference((int column, int row) pair)
        => new((uint)pair.column, (uint)pair.row);

    private static void CheckValidCell(uint column, uint row)
    {
        // Excel limitations 1,048,576 rows by 16,384 columns
        // https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3

        if (column is < 1 or > 16_384)
            throw new ArgumentOutOfRangeException(nameof(column), "Column should be in range [1 .. 16,384].");

        if (row is < 1 or > 1_048_576)
            throw new ArgumentOutOfRangeException(nameof(row), "Row should be in range [1 .. 1,048,576].");
    }
}
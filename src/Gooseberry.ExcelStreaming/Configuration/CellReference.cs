using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Configuration
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct CellReference
    {
        internal static readonly char[] ColumnAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToArray();
        internal static readonly char[] RowAlphabet = "0123456789".ToArray();
        
        public static CellReference Parse(string alias)
        {
            var (row, column) = ParseAlias(alias);
            return new CellReference(row, column);
        }
        
        public CellReference(uint row, uint column)
        {
            EnsureCellPosition(row, column);
            Row = row;
            Column = column;
        }

        internal readonly uint Column;
        
        internal readonly uint Row;
        
        private static void EnsureCellPosition(uint row, uint column)
        {
            // Excel limitations 1,048,576 rows by 16,384 columns
            // https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
            
            if (column is < 1 or > 16_384)
                throw new ArgumentOutOfRangeException(nameof(column), "Column should be in range [1 .. 16,384].");
         
            if (row is < 1 or > 1_048_576)
                throw new ArgumentOutOfRangeException(nameof(row), "Row should be in range [1 .. 1,048,576].");
        }
        
        private static (uint row, uint column) ParseAlias(string alias)
        {
            var (row, position) = Parse(alias, position: alias.Length - 1, RowAlphabet);
            var (column, _) = Parse(alias, position, ColumnAlphabet);            

            // Column alias A is 1, so return column + 1 
            return (row, column + 1);
        }
        
        private static (uint value, int position) Parse(string alias, int position, char[] alphabet)
        {
            uint @base = 1; 
            uint value = 0;
            do
            {
                var digit = Array.IndexOf(alphabet, alias[position]);
                if (digit < 0)
                    break;

                value += @base * (uint)digit;
                @base *= (uint)alphabet.Length;
                position--;
            } while (position > 0);

            return (value, position);
        }        
    }
}
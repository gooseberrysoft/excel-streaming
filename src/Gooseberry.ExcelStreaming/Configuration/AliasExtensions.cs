using System;
using System.Linq;

namespace Gooseberry.ExcelStreaming.Configuration
{
    internal static class AliasExtensions
    {
        // Excel limitations 1,048,576 rows by 16,384 columns
        // https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
        
        private static readonly char[] ColumnAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToArray();
        private static readonly char[] RowAlphabet = "0123456789".ToArray();

        public static int FormatColumnAlias(this uint column, Span<char> destination)
        {
            if (column is < 1 or > 16_384)
                throw new ArgumentOutOfRangeException(nameof(column), "Column should be in range [1 .. 16,384].");

            // Column alias A is 1, so pass column - 1 
            return Convert(column - 1, ColumnAlphabet, destination);
        }
        
        public static int FormatRowAlias(this uint row, Span<char> destination)
        {
            if (row is < 1 or > 1_048_576)
                throw new ArgumentOutOfRangeException(nameof(row), "Row should be in range [1 .. 1,048,576].");

            return Convert(row, RowAlphabet, destination);
        }

        public static (uint row, uint column) Parse(this string alias)
        {
            var (row, position) = Parse(alias, position: alias.Length - 1, RowAlphabet);
            var (column, _) = Parse(alias, position, ColumnAlphabet);            

            // Column alias A is 1, so pass column + 1 
            return (row, column + 1);
        }
        
        private static int Convert(uint value, char[] alphabet, Span<char> destination)
        {
            var destinationIndex = destination.Length - 1;
            
            var remain = (int)value;
            do
            {
                var charIndex = remain % alphabet.Length;

                destination[destinationIndex] = alphabet[charIndex];

                destinationIndex--;
                remain /= alphabet.Length;
            } while (remain > 0);

            return destination.Length - destinationIndex - 1;
        }

        public static (uint value, int position) Parse(string alias, int position, char[] alphabet)
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
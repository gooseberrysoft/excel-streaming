using System;
using System.Linq;

namespace Gooseberry.ExcelStreaming
{
    internal static class AliasExtensions
    {
        internal static readonly char[] ColumnAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToArray();
        internal static readonly char[] RowAlphabet = "0123456789".ToArray();

        public static (uint row, uint column) Parse(this string alias)
        {
            var (row, position) = Parse(alias, position: alias.Length - 1, RowAlphabet);
            var (column, _) = Parse(alias, position, ColumnAlphabet);            

            // Column alias A is 1, so return column + 1 
            return (row, column + 1);
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
using System;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Configuration
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct CellReference
    {
        public CellReference(string alias)
        {
            var (row, column) = alias.Parse();
            Row = row;
            Column = column;
            Alias = alias;
        }
        
        public CellReference(uint row, uint column)
        {
            Row = row;
            Column = column;
            Alias = GetAlias(row, column);
        }

        public uint Column { get; }
        
        public uint Row { get; }
        
        public string Alias { get; }
        
        private static string GetAlias(uint row, uint column)
        {
            Span<char> chars = stackalloc char[10];

            var written = row.FormatRowAlias(chars);
            written += column.FormatColumnAlias(chars[..^written]);

            return chars[^written..].ToString();
        }
    }
}
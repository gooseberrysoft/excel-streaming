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
        }
        
        public CellReference(uint row, uint column)
        {
            Row = row;
            Column = column;
        }

        internal readonly uint Column;
        
        internal readonly uint Row;
        
        internal void WriteAlias(BufferedWriter writer)
        {
            Span<char> chars = stackalloc char[10];

            var written = Row.FormatRowAlias(chars);
            written += Column.FormatColumnAlias(chars[..^written]);

            writer.Write(chars[^written..]);
        }
    }
}
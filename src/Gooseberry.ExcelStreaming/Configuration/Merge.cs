using System;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Configuration
{
    [StructLayout(LayoutKind.Auto)]
    internal readonly struct Merge
    {
        private readonly uint _fromRow;
        private readonly uint _fromColumn;
        private readonly uint _downSize;
        private readonly uint _rightSize;

        public Merge(uint fromRow, uint fromColumn, uint downSize, uint rightSize)
        {
            _fromRow = fromRow;
            _fromColumn = fromColumn;
            _downSize = downSize;
            _rightSize = rightSize;
        }

        public void WriteAlias(BufferedWriter writer)
        {
            Span<char> chars = stackalloc char[21];

            var bottomRow = _fromRow + _downSize;
            var rightColumn = _fromColumn + _rightSize;
            
            var written = bottomRow.FormatRowAlias(chars);
            written += rightColumn.FormatColumnAlias(chars[..^written]);

            written += 1;
            chars[^written] = ':';
            
            written += _fromRow.FormatRowAlias(chars[..^written]);
            written += _fromColumn.FormatColumnAlias(chars[..^written]);  
            
            writer.Write(chars[^written..]);
        }
    }
}
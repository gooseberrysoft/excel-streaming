using System;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Configuration
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct Merge
    {
        public Merge(uint fromRow, uint fromColumn, uint downSize, uint rightSize)
        {
            Span<char> chars = stackalloc char[21];

            var bottomRow = fromRow + downSize;
            var rightColumn = fromColumn + rightSize;
            
            var written = bottomRow.FormatRowAlias(chars);
            written += rightColumn.FormatColumnAlias(chars[..^written]);

            written += 1;
            chars[^written] = ':';
            
            written += fromRow.FormatRowAlias(chars[..^written]);
            written += fromColumn.FormatColumnAlias(chars[..^written]);

            Alias = chars[^written..].ToString();
        }

        public string Alias { get; }
    }
}
using System;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Configuration
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct Column
    {
        public Column(decimal width)
        {
            if (width <= 0)
                throw new ArgumentException("Column width cannot be less or equal zero.", nameof(width));

            Width = width;
        }

        public decimal Width { get; }
    }
}
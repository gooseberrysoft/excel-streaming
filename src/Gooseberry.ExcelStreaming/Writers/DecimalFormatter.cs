using System;
using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct DecimalFormatter : IValueFormatter<decimal>
{
    private static readonly int MaxChars = decimal.MinValue.ToString().Length + 2;

    public int MaximumChars => MaxChars;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool TryFormat(in decimal value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes);
}
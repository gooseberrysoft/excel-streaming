using System;
using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct IntFormatter : INumberFormatter<int>
{
    private static readonly int MaxChars = int.MinValue.ToString().Length;

    public int MaximumChars => MaxChars;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool TryFormat(in int value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes);
}
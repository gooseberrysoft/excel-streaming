using System;
using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class IntWriter : ValueWriter<int>
{
    private static readonly int MaxChars = int.MinValue.ToString().Length;

    protected override int MaximumChars => MaxChars;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    protected override bool TryFormat(in int value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes);
}
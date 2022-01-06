using System;
using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class LongWriter : ValueWriter<long>
{
    private static readonly int MaxChars = long.MinValue.ToString().Length;

    protected override int MaximumChars => MaxChars;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    protected override bool TryFormat(in long value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes);
}
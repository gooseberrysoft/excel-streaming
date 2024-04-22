using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct LongFormatter : INumberFormatter<long>
{
    private static readonly int MaxChars = long.MinValue.ToString().Length;

    public int MaximumChars => MaxChars;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool TryFormat(in long value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes);
}
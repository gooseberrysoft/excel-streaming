using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct UIntFormatter : INumberFormatter<uint>
{
    private static readonly int MaxChars = int.MinValue.ToString().Length;

    public int MaximumChars => MaxChars;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool TryFormat(in uint value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes);
}
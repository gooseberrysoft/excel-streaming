using System.Buffers;
using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct IntHex8Formatter : INumberFormatter<int>
{
    public int MaximumChars => 8;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool TryFormat(in int value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes, new StandardFormat('x', 8));
}
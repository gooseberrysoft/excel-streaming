using System.Buffers;
using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers.Formatters;

internal class IntHex8Formatter : INumberFormatter<int>
{
    public static int MaximumChars => 8;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryFormat(in int value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes, new StandardFormat('x', 8));
}
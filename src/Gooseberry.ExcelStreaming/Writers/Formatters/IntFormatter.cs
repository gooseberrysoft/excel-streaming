using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers.Formatters;

internal class IntFormatter : INumberFormatter<int>
{
    public static int MaximumChars { get; } = int.MinValue.ToString().Length;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryFormat(in int value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes);
}
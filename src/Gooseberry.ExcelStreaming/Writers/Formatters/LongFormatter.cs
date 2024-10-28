using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers.Formatters;

internal class LongFormatter : INumberFormatter<long>
{
    public static int MaximumChars { get; } = long.MinValue.ToString().Length;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryFormat(in long value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes);
}
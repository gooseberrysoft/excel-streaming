using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers.Formatters;

internal class DecimalFormatter : INumberFormatter<decimal>
{
    public static int MaximumChars { get; } = decimal.MinValue.ToString().Length + 2;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryFormat(in decimal value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes);
}
using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers.Formatters;

internal class DoubleFormatter : INumberFormatter<double>
{
    public static int MaximumChars => 26;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryFormat(in double value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes);
}
using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers.Formatters;

internal class UIntFormatter : INumberFormatter<uint>
{
    public static int MaximumChars { get; } = int.MinValue.ToString().Length;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryFormat(in uint value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value, destination, out encodedBytes);
}
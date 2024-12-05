using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers.Formatters;

internal class DateOnlyFormatter : INumberFormatter<DateOnly>
{
    public static int MaximumChars { get; } = double.MinValue.ToString().Length + 2;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryFormat(in DateOnly value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value.ToDateTime(default).ToOADate(), destination, out encodedBytes);
}
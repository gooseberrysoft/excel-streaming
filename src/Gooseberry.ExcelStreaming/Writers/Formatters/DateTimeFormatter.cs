using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers.Formatters;

internal class DateTimeFormatter : INumberFormatter<DateTime>
{
    public static int MaximumChars { get; } = double.MinValue.ToString().Length + 2;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryFormat(in DateTime value, Span<byte> destination, out int encodedBytes)
        => Utf8Formatter.TryFormat(value.ToOADate(), destination, out encodedBytes);
}
using System.Buffers.Text;
using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct DateTimeFormatter : INumberFormatter<DateTime>
{
    private static readonly int MaxChars = double.MinValue.ToString().Length + 2;

    public int MaximumChars => MaxChars;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool TryFormat(in DateTime value, Span<byte> destination, out int encodedBytes) 
        => Utf8Formatter.TryFormat(value.ToOADate(), destination, out encodedBytes);
}
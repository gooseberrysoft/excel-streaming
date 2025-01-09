#if NET8_0_OR_GREATER
using System.Buffers.Text;
using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Styles;
using System.Runtime.CompilerServices;
using System.Text;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class Utf8DateTimeCellWriter
{
    private static readonly int NumberSize = long.MinValue.ToString().Length + 5;

    private static ReadOnlySpan<byte> StylePrefix => "<c s=\""u8;
    private static ReadOnlySpan<byte> StylePostfix => "\"><v>"u8;
    private static ReadOnlySpan<byte> Postfix => "</v></c>"u8;

    private static int DefaultStyleValue = StylesSheetBuilder.Default.DefaultDateStyle.Value;
    private static byte[] DefaultStyle = StylePrefix.Combine(Encoding.UTF8.GetBytes(DefaultStyleValue.ToString()), StylePostfix);

    private static readonly int Size = StylePrefix.Length + NumberSize + StylePostfix.Length
        + NumberSize
        + Postfix.Length;

    public static void Write(DateTime value, BuffersChain buffer, StyleReference style)
    {
        var span = buffer.GetSpan(Size);
        var written = 0;

        if (style.Value == DefaultStyleValue)
        {
            DefaultStyle.CopyTo(span);
            span = span.Slice(DefaultStyle.Length);
            written = DefaultStyle.Length;
        }
        else
        {
            StylePrefix.CopyTo(ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(style.Value, buffer, ref span, ref written);
            StylePostfix.CopyTo(ref span, ref written);
        }

        FormatOADate(value, span, out var writtenBytes);

        span = span.Slice(writtenBytes);
        written += writtenBytes;

        Postfix.WriteAdvanceTo(buffer, span, written);
    }

    public static void Write(DateOnly value, BuffersChain buffer, StyleReference style)
    {
        var span = buffer.GetSpan(Size);
        var written = 0;

        if (style.Value == DefaultStyleValue)
        {
            DefaultStyle.CopyTo(span);
            span = span.Slice(DefaultStyle.Length);
            written = DefaultStyle.Length;
        }
        else
        {
            StylePrefix.CopyTo(ref span, ref written);
            Utf8SpanFormattableWriter.WriteValue(style.Value, buffer, ref span, ref written);
            StylePostfix.CopyTo(ref span, ref written);
        }

        Utf8SpanFormattableWriter.WriteValue(value.ToOADate(), buffer, ref span, ref written);

        Postfix.WriteAdvanceTo(buffer, span, written);
    }


    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void FormatOADate(DateTime value, Span<byte> span, out int written)
    {
        written = 0;

        //see DateTime.ToOADate()
        long ticks = value.Ticks - DoubleDateOffset;
        if (ticks < 0)
        {
            long frac = ticks % TicksPerDay;
            if (frac != 0) ticks -= (TicksPerDay + frac) * 2;

            span[0] = (byte)'-';
            ++written;

            ticks = -ticks;
        }

        //ticks positive 

        var gigaOADate = ticks / GigaTicksPerDay;
        var days = (int)(gigaOADate / Giga);
        var time = gigaOADate - Giga * days;

        if (!days.TryFormat(span.Slice(written), out var daysBytesWritten))
            ThrowNotEnoughMemory();

        written += daysBytesWritten;

        if (time == 0)
            return;

        span[written] = (byte)'.';

        if (!time.TryFormat(span.Slice(written + 1), out var timeBytesWritten))
            ThrowNotEnoughMemory();

        written += 1 + timeBytesWritten;
    }

    private static void ThrowNotEnoughMemory()
        => throw new InvalidOperationException(
            "Not enough span length to format datetime. It should never happen. Please create issue on github.");

    /*** From DateTime source code ***/

    // Number of 100ns ticks per time unit
    internal const int MicrosecondsPerMillisecond = 1000;
    private const long TicksPerMicrosecond = 10;
    private const long TicksPerMillisecond = TicksPerMicrosecond * MicrosecondsPerMillisecond;

    private const int HoursPerDay = 24;
    private const long TicksPerSecond = TicksPerMillisecond * 1000;
    private const long TicksPerMinute = TicksPerSecond * 60;
    private const long TicksPerHour = TicksPerMinute * 60;
    private const long TicksPerDay = TicksPerHour * HoursPerDay; // 864_000_000_000	long

    private const long GigaTicksPerDay = TicksPerDay / Giga; // 864	long
    private const long Giga = 1_000_000_000L;


    // Number of days in a non-leap year
    private const int DaysPerYear = 365;

    // Number of days in 4 years
    private const int DaysPer4Years = DaysPerYear * 4 + 1; // 1461

    // Number of days in 100 years
    private const int DaysPer100Years = DaysPer4Years * 25 - 1; // 36524

    // Number of days in 400 years
    private const int DaysPer400Years = DaysPer100Years * 4 + 1; // 146097

    // Number of days from 1/1/0001 to 12/30/1899
    private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;

    private const long DoubleDateOffset = DaysTo1899 * TicksPerDay;
}
#endif
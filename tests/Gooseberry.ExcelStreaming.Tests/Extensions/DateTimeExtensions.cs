using System.Text;
using Gooseberry.ExcelStreaming.Writers.Cells;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming.Tests;

public static class DateTimeExtensions
{
    public static string ToInternalOADate(this DateTime value)
    {
        Span<byte> span = stackalloc byte[Utf8DateTimeCellWriter.NumberSize];

        Utf8DateTimeCellWriter.FormatOADate(value, span, out var written);

        return Encoding.UTF8.GetString(span.Slice(0, written));
    }
}
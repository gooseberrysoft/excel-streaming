#if NET8_0_OR_GREATER
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.Benchmarks;

/*
| Method              | Mean       | Error     | StdDev    |
|-------------------- |-----------:|----------:|----------:|
| DateCustomFormatter |   5.815 ms | 0.0754 ms | 0.0706 ms |
| CustomFormatter     |  13.545 ms | 0.1515 ms | 0.1343 ms |
| DateDoubleFormatter |  76.139 ms | 1.0479 ms | 0.9289 ms |
| DoubleFormatter     | 113.291 ms | 1.5578 ms | 1.4572 ms |
*/

[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class OADateBenchmarks
{
    private static readonly DateTime Value = DateTime.Now;
    private static readonly DateTime DateValue = DateTime.Now.Date;

    [Benchmark]
    public int CustomFormatter()
    {
        var bytes = new byte[256];
        int total = 0;

        for (int i = 0; i < 1_000_000; ++i)
        {
            Utf8DateTimeCellWriter.FormatOADate(Value, bytes, out var written);
            total += written;
        }

        return total;
    }

    [Benchmark]
    public int DoubleFormatter()
    {
        var bytes = new byte[256];
        int total = 0;

        for (int i = 0; i < 1_000_000; ++i)
        {
            Value.ToOADate().TryFormat(bytes, out var written);

            total += written;
        }

        return total;
    }

    [Benchmark]
    public int DateCustomFormatter()
    {
        var bytes = new byte[256];
        int total = 0;

        for (int i = 0; i < 1_000_000; ++i)
        {
            Utf8DateTimeCellWriter.FormatOADate(DateValue, bytes, out var written);
            total += written;
        }

        return total;
    }

    [Benchmark]
    public int DateDoubleFormatter()
    {
        var bytes = new byte[256];
        int total = 0;

        for (int i = 0; i < 1_000_000; ++i)
        {
            DateValue.ToOADate().TryFormat(bytes, out var written);

            total += written;
        }

        return total;
    }
}
#endif
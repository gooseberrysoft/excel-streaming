#if NET8_0_OR_GREATER
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.Benchmarks;

/*
| Method                 | Mean      | Error    | StdDev   |
|----------------------- |----------:|---------:|---------:|
| CustomFormatterHandler |  13.48 ms | 0.113 ms | 0.101 ms |
| DoubleFormatter        | 116.92 ms | 1.733 ms | 1.621 ms |
*/

[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class OADateBenchmarks
{
    private static readonly DateTime Value = DateTime.Now;

    [Benchmark]
    public int CustomFormatterHandler()
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
}
#endif
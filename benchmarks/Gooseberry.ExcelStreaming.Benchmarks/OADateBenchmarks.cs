#if NET8_0_OR_GREATER
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.Benchmarks;

/*
| Method              | Mean       | Error     | StdDev     | Median     |
|-------------------- |-----------:|----------:|-----------:|-----------:|
| DateCustomFormatter |   3.574 ms | 0.1830 ms |  0.5396 ms |   3.413 ms |
| CustomFormatter     |  11.247 ms | 0.3841 ms |  1.1143 ms |  11.017 ms |
| DateDoubleFormatter |  83.862 ms | 1.6674 ms |  1.7123 ms |  83.485 ms |
| DoubleFormatter     | 144.717 ms | 5.4055 ms | 15.6822 ms | 142.020 ms |
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
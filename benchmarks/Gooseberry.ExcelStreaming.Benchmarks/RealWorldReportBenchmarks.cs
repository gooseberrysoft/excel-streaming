using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;

namespace Gooseberry.ExcelStreaming.Benchmarks;

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class RealWorldReportBenchmarks
{
    [Params(100, 1000, 10_000, 100_000, 500_000)]
    public int RowsCount { get; set; }

    [Benchmark]
    public async Task RealWorldReport()
    {
        await using var writer = new ExcelWriter(Stream.Null);

        await writer.StartSheet("PNL");
        var dateTime = new DateTime(638721164006405476);

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (int i = 0; i < 5; i++)
            {
                writer.AddCell(row);
                writer.AddCell(DateOnly.FromDateTime(dateTime));
                writer.AddCell("\"Alice’s Adventures in Wonderland\" by Lewis Carroll");
                writer.AddCell(1789);
                writer.AddCell(1234567.9876M);
                writer.AddCell(-936.9M);
                writer.AddCell(0.999M);
                writer.AddCell(23.00M);
                writer.AddCell(56688900.56M);
                writer.AddCell("7895-654-098-45");
                writer.AddCell(1789);
                writer.AddCell(1234567.9876M);
                writer.AddCell(-936.9M);
                writer.AddCell(0.999M);
                writer.AddCell(23.00M);
                writer.AddCell(56688900.56M);
                writer.AddCell(784000);
                writer.AddCell(dateTime);
                writer.AddCell(56688900.56M);
                writer.AddCell(784000);
            }
        }

        await writer.Complete();
    }
}
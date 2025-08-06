using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;

namespace Gooseberry.ExcelStreaming.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
[SimpleJob(RuntimeMoniker.Net90)]
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
                writer.AddCell(row)
                    .AddCell(DateOnly.FromDateTime(dateTime))
                    .AddCell("\"Alice’s Adventures in Wonderland\" by Lewis Carroll")
                    .AddCell(1789)
                    .AddCell(1234567.9876M)
                    .AddCell(-936.9M)
                    .AddCell(0.999M)
                    .AddCell(23.00M)
                    .AddCell(56688900.56M)
                    .AddCell("7895-654-098-45")
                    .AddCell(1789)
                    .AddCell(1234567.9876M)
                    .AddCell(-936.9M)
                    .AddCell(0.999M)
                    .AddCell(23.00M)
                    .AddCell(56688900.56M)
                    .AddCell(784000)
                    .AddCell(dateTime)
                    .AddCell(56688900.56M)
                    .AddCell(784000);
            }
        }

        await writer.Complete();
    }

    private sealed class NullZip : IZipArchive
    {
        public void Dispose()
        {
        }

        public Stream CreateEntry(string entryPath)
            => Stream.Null;
    }
}
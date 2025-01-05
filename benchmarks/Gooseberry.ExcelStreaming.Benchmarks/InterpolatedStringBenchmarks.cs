#if NET8_0_OR_GREATER
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;

namespace Gooseberry.ExcelStreaming.Benchmarks;

/*
   | Method                    | Mean     | Error    | StdDev   | Gen0       | Allocated   |
   |-------------------------- |---------:|---------:|---------:|-----------:|------------:|
   | InterpolatedStringHandler | 684.9 ms | 13.39 ms | 14.32 ms |          - |      7.5 KB |
   | RegularString             | 959.5 ms | 18.83 ms | 25.78 ms | 81000.0000 | 500007.3 KB | 
 */

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class InterpolatedStringBenchmarks
{
    private const int ColumnBatchesCount = 10;
    private const int RowsCount = 100_000;
    public static string Hampty = "Humpty Dumpty";
    public static string Alice = "Alice";

    [Benchmark]
    public async Task InterpolatedStringHandler()
    {
        await using var writer = new ExcelWriter(new NullZipArchive());

        await writer.StartSheet("test");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (var columnBatch = 0; columnBatch < ColumnBatchesCount; columnBatch++)
            {
                writer.AddCell($"Aligned row number: {row,10}, at {DateTime.Now:D}");
                writer.AddCell($"<Random> value: {Guid.NewGuid():B} with B format");
                writer.AddCell($"'It's very provoking,' {Hampty} said after a long silence, looking away from {Alice} as he spoke");
            }
        }

        await writer.Complete();
    }

    [Benchmark]
    public async Task RegularString()
    {
        await using var writer = new ExcelWriter(new NullZipArchive());

        await writer.StartSheet("test");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (var columnBatch = 0; columnBatch < ColumnBatchesCount; columnBatch++)
            {
                var s1 = $"Aligned row number: {row,10}, at {DateTime.Now:D}";
                var s2 = $"<Random> value: {Guid.NewGuid():B} with B format";
                var s3 = $"'It's very provoking,' {Hampty} said after a long silence, looking away from {Alice} as he spoke";

                writer.AddCell(s1);
                writer.AddCell(s2);
                writer.AddCell(s3);
            }
        }

        await writer.Complete();
    }

    public sealed class NullZipArchive : IZipArchive
    {
        public void Dispose()
        {
        }

        public Stream CreateEntry(string entryPath)
        {
            return Stream.Null;
        }
    }
}
#endif
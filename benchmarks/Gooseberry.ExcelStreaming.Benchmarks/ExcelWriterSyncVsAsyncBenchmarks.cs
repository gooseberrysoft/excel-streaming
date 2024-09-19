using System.IO.Compression;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;

namespace Gooseberry.ExcelStreaming.Benchmarks;

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class ExcelWriterSyncVsAsyncBenchmarks
{
    private const int ColumnBatchesCount = 10;

    [Params(100, 1000, 10_000, 100_000)]
    public int RowsCount { get; set; }

    [Benchmark]
    public async Task AsyncZip()
    {
        await using var writer = new ExcelWriter(new DefaultZipArchive(Stream.Null, CompressionLevel.Optimal), async: true);

        await writer.StartSheet("test");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (var columnBatch = 0; columnBatch < ColumnBatchesCount; columnBatch++)
            {
                writer.AddCell(row);
                writer.AddCell(DateTime.Now.Ticks);
                writer.AddCell(DateTime.Now);
                writer.AddCell(1234567.9876M);
                writer.AddCell("Tags such as <img> and <input> directly introduce content into the page.");
                writer.AddCell("The cat (Felis catus), commonly referred to as the domestic cat");
                writer.AddCellSharedString(
                    "The dog (Canis familiaris or Canis lupus familiaris) is a domesticated descendant of the wolf");
            }
        }

        await writer.Complete();
    }


    [Benchmark]
    public async Task SyncZip()
    {
        await using var writer = new ExcelWriter(new DefaultZipArchive(Stream.Null, CompressionLevel.Optimal), async: false);

        await writer.StartSheet("test");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (var columnBatch = 0; columnBatch < ColumnBatchesCount; columnBatch++)
            {
                writer.AddCell(row);
                writer.AddCell(DateTime.Now.Ticks);
                writer.AddCell(DateTime.Now);
                writer.AddCell(1234567.9876M);
                writer.AddCell("Tags such as <img> and <input> directly introduce content into the page.");
                writer.AddCell("The cat (Felis catus), commonly referred to as the domestic cat");
                writer.AddCellSharedString(
                    "The dog (Canis familiaris or Canis lupus familiaris) is a domesticated descendant of the wolf");
            }
        }

        await writer.Complete();
    }
}
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Order;

namespace Gooseberry.ExcelStreaming.Benchmarks;

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
//[Config(typeof(Config))]
public class ExcelWriterVersionsBenchmarks
{
    private class Config : ManualConfig
    {
        public Config()
        {
            AddJob(Job.Default.WithNuGet(new NuGetReferenceList()
            {
                new NuGetReference("Gooseberry.ExcelStreaming", "1.3.0"),
            }));

            AddJob(Job.Default.WithNuGet(new NuGetReferenceList()
            {
                new NuGetReference("Gooseberry.ExcelStreaming", "1.5.0"),
            }));
        }
    }

    [Params(100, 1000, 10_000, 100_000, 500_000)]
    public int RowsCount { get; set; }

    private const int ColumnBatchesCount = 10;

    [Benchmark]
    public async Task ExcelWriter_WithEscaping()
    {
        await using var writer = new ExcelWriter(Stream.Null);

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
                writer.AddCellWithSharedString(
                    "The dog (Canis familiaris or Canis lupus familiaris) is a domesticated descendant of the wolf");
            }
        }

        await writer.Complete();
    }

    [Benchmark]
    public async Task ExcelWriter_WithoutEscaping()
    {
        await using var writer = new ExcelWriter(Stream.Null);

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
                writer.AddCell("Tags such as _img_ and _input_ directly introduce content into the page.");
                writer.AddCell("The cat (Felis catus), commonly referred to as the domestic cat");
                writer.AddCellWithSharedString(
                    "The dog (Canis familiaris or Canis lupus familiaris) is a domesticated descendant of the wolf");
            }
        }

        await writer.Complete();
    }

    [Benchmark]
    public async Task ExcelWriter_WithEscaping_Utf8()
    {
        await using var writer = new ExcelWriter(Stream.Null);

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
                writer.AddUtf8Cell("Tags such as <img> and <input> directly introduce content into the page."u8);
                writer.AddUtf8Cell("The cat (Felis catus), commonly referred to as the domestic cat"u8);
                writer.AddCellWithSharedString(
                    "The dog (Canis familiaris or Canis lupus familiaris) is a domesticated descendant of the wolf");
            }
        }

        await writer.Complete();
    }

    [Benchmark]
    public async Task ExcelWriter_WithoutEscaping_Utf8()
    {
        await using var writer = new ExcelWriter(Stream.Null);

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
                writer.AddUtf8Cell("Tags such as _img_ and _input_ directly introduce content into the page."u8);
                writer.AddUtf8Cell("The cat (Felis catus), commonly referred to as the domestic cat"u8);
                writer.AddCellWithSharedString(
                    "The dog (Canis familiaris or Canis lupus familiaris) is a domesticated descendant of the wolf");
            }
        }

        await writer.Complete();
    }
}
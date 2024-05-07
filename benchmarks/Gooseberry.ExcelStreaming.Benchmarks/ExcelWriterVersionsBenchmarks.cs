using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Order;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;

namespace Gooseberry.ExcelStreaming.Benchmarks;

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
[Config(typeof(Config))]
public class ExcelWriterVersionsBenchmarks
{
    [Params(100, 1000, 10_000, 100_000, 500_000)]
    public int RowsCount { get; set; }

    private const int ColumnBatchesCount = 10;

    [Benchmark]
    public async Task ExcelWriter_RealWorldReport()
    {
        await using var writer = new ExcelWriter(Stream.Null);

        await writer.StartSheet("PNL");
        var dateTime = DateTime.Now;

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            writer.AddCell(row);
            writer.AddCell(dateTime);
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

        await writer.Complete();
    }

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

    private sealed class Config : ManualConfig
    {
        public Config()
            => AddColumn(new VersionColumn("1.5.1"));
    }

    public sealed class VersionColumn(string value) : IColumn
    {
        public string Id => nameof(VersionColumn);

        public string ColumnName => "Version";

        public string Legend => "Version";

        public UnitType UnitType => UnitType.Size;

        public bool AlwaysShow => true;

        public ColumnCategory Category => ColumnCategory.Meta;

        public int PriorityInCategory => 0;

        public bool IsNumeric => false;

        public bool IsAvailable(Summary summary) => true;

        public bool IsDefault(Summary summary, BenchmarkCase benchmarkCase) => false;

        public string GetValue(Summary summary, BenchmarkCase benchmarkCase) => GetValue(summary, benchmarkCase, SummaryStyle.Default);

        public string GetValue(Summary summary, BenchmarkCase benchmarkCase, SummaryStyle style) => value;

        public override string ToString() => ColumnName;
    }
}
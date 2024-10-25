#if NET8_0_OR_GREATER
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;

namespace Gooseberry.ExcelStreaming.Benchmarks;

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class Utf8FormattableBenchmarks
{
    private const int ColumnBatchesCount = 10;
    private const int RowsCount = 100_000;

    [Benchmark]
    public async Task Utf8FormattableNumber()
    {
        await using var writer = new ExcelWriter(new NullZipArchive());

        await writer.StartSheet("test");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (var columnBatch = 0; columnBatch < ColumnBatchesCount; columnBatch++)
            {
                writer.AddCellNumber(row);
                writer.AddCellNumber(DateTime.Now.Ticks);
                writer.AddCellNumber(234.897M);
                writer.AddCellNumber(1234567.9876M);
            }
        }

        await writer.Complete();
    }

    [Benchmark]
    public async Task FormatterNumber()
    {
        await using var writer = new ExcelWriter(new NullZipArchive());

        await writer.StartSheet("test");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (var columnBatch = 0; columnBatch < ColumnBatchesCount; columnBatch++)
            {
                writer.AddCell(row);
                writer.AddCell(DateTime.Now.Ticks);
                writer.AddCell(234.897M);
                writer.AddCell(1234567.9876M);
            }
        }

        await writer.Complete();
    }

    [Benchmark]
    public async Task Utf8FormattableString()
    {
        await using var writer = new ExcelWriter(new NullZipArchive());

        await writer.StartSheet("test");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (var columnBatch = 0; columnBatch < ColumnBatchesCount; columnBatch++)
            {
                writer.AddCellString(row);
                writer.AddCellString(DateTime.Now);
                writer.AddCellString(Guid.NewGuid());
                writer.AddCellString(1234567.9876M);
                writer.AddCellString(DateOnly.FromDateTime(DateTime.Now));
            }
        }

        await writer.Complete();
    }

    [Benchmark]
    public async Task EncoderString()
    {
        await using var writer = new ExcelWriter(new NullZipArchive());

        await writer.StartSheet("test");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (var columnBatch = 0; columnBatch < ColumnBatchesCount; columnBatch++)
            {
                writer.AddCell(row.ToString());
                writer.AddCell(DateTime.Now.ToString());
                writer.AddCell(Guid.NewGuid().ToString());
                writer.AddCell(1234567.9876M.ToString());
                writer.AddCell(DateOnly.FromDateTime(DateTime.Now).ToString());
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
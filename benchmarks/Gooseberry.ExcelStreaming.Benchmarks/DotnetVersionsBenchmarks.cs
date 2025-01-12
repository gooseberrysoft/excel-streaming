using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Order;

namespace Gooseberry.ExcelStreaming.Benchmarks;

[SimpleJob(RuntimeMoniker.Net80)]
[SimpleJob(RuntimeMoniker.Net90)]
public class DotnetVersionsBenchmarks
{
    private string[] _simpleStrings = null!;
    private string[] _escapingStrings = null!;

    [GlobalSetup]
    public void GlobalSetup()
    {
        _simpleStrings = Enumerable.Range(0, RowsCount * 5)
            .Select(i => $"row col {i} text")
            .ToArray();

        _escapingStrings = Enumerable.Range(0, RowsCount * 5)
            .Select(i => $"row col {i} text with <tag> & \"quote\"'s")
            .ToArray();
    }

    [Params(100_000)]
    public int RowsCount { get; set; }

    [Benchmark]
    public async Task SimpleStringsWriting()
    {
        await using var writer = new ExcelWriter(Stream.Null);

        await writer.StartSheet("PNL");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (int i = 0; i < 5; i++)
            {
                writer.AddCell(_simpleStrings[row * 5 + i]);
            }
        }

        await writer.Complete();
    }

    [Benchmark]
    public async Task EscapedStringsWriting()
    {
        await using var writer = new ExcelWriter(Stream.Null);

        await writer.StartSheet("PNL");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (int i = 0; i < 5; i++)
            {
                writer.AddCell(_escapingStrings[row * 5 + i]);
            }
        }

        await writer.Complete();
    }

    [Benchmark]
    public async Task DateTimeWriting()
    {
        await using var writer = new ExcelWriter(Stream.Null);

        await writer.StartSheet("PNL");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (int i = 0; i < 5; i++)
            {
                writer.AddCell(DateTime.Now);
            }
        }

        await writer.Complete();
    }

    [Benchmark]
    public async Task IntWriting()
    {
        await using var writer = new ExcelWriter(Stream.Null);

        await writer.StartSheet("PNL");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (int i = 0; i < 5; i++)
            {
                writer.AddCell(row * i);
            }
        }

        await writer.Complete();
    }

    [Benchmark]
    public async Task DecimalWriting()
    {
        await using var writer = new ExcelWriter(Stream.Null);

        await writer.StartSheet("PNL");

        for (var row = 0; row < RowsCount; row++)
        {
            await writer.StartRow();

            for (int i = 0; i < 5; i++)
            {
                writer.AddCell(1_345_767_874.56789M);
            }
        }

        await writer.Complete();
    }
}
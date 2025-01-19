using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Order;

namespace Gooseberry.ExcelStreaming.Benchmarks;

[SimpleJob(RuntimeMoniker.Net80)]
[SimpleJob(RuntimeMoniker.Net90)]
[Orderer(SummaryOrderPolicy.Method)]
public class DotnetVersionsBenchmarks
{
    private string[] _simpleStrings = null!;
    private string[] _escapingStrings = null!;
    private DateTime[] _dates = null!;
    private int[] _ints = null!;
    private decimal[] _decimals = null!;

    [GlobalSetup]
    public void GlobalSetup()
    {
        _simpleStrings = Enumerable.Range(0, RowsCount * 5)
            .Select(i => $"row col {i} text")
            .ToArray();

        _escapingStrings = Enumerable.Range(0, RowsCount * 5)
            .Select(i => $"row col {i} text with <tag> & \"quote\"'s")
            .ToArray();

        _dates = Enumerable.Range(0, RowsCount * 5)
            .Select(_ => DateTime.Now.AddTicks(Random.Shared.Next()))
            .ToArray();

        _ints = Enumerable.Range(0, RowsCount * 5)
            .Select(i => i + 123_456_789)
            .ToArray();

        _decimals = Enumerable.Range(0, RowsCount * 5)
            .Select(i => 1_345_767_874.56789M + i * 123.76M)
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

            int i = 0;
            writer.AddCell(_simpleStrings[row * 5 + i++]);
            writer.AddCell(_simpleStrings[row * 5 + i++]);
            writer.AddCell(_simpleStrings[row * 5 + i++]);
            writer.AddCell(_simpleStrings[row * 5 + i++]);
            writer.AddCell(_simpleStrings[row * 5 + i++]);
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
            int i = 0;
            writer.AddCell(_escapingStrings[row * 5 + i++]);
            writer.AddCell(_escapingStrings[row * 5 + i++]);
            writer.AddCell(_escapingStrings[row * 5 + i++]);
            writer.AddCell(_escapingStrings[row * 5 + i++]);
            writer.AddCell(_escapingStrings[row * 5 + i++]);
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

            int i = 0;
            writer.AddCell(_dates[row * 5 + i++]);
            writer.AddCell(_dates[row * 5 + i++]);
            writer.AddCell(_dates[row * 5 + i++]);
            writer.AddCell(_dates[row * 5 + i++]);
            writer.AddCell(_dates[row * 5 + i++]);
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

            int i = 0;
            writer.AddCell(_ints[row * 5 + i++]);
            writer.AddCell(_ints[row * 5 + i++]);
            writer.AddCell(_ints[row * 5 + i++]);
            writer.AddCell(_ints[row * 5 + i++]);
            writer.AddCell(_ints[row * 5 + i++]);
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

            int i = 0;
            writer.AddCell(_decimals[row * 5 + i++]);
            writer.AddCell(_decimals[row * 5 + i++]);
            writer.AddCell(_decimals[row * 5 + i++]);
            writer.AddCell(_decimals[row * 5 + i++]);
            writer.AddCell(_decimals[row * 5 + i++]);
        }

        await writer.Complete();
    }
}
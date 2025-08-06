using FluentAssertions;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class BufferTests
{
    [Fact]
    public void InitBuffer_CorrectProperties()
    {
        var size = 64;
        var buffer = new Buffer(size, new());

        buffer.RemainingCapacity.Should().Be(size);
        buffer.Written.Should().Be(0);
        buffer.GetSpan().Length.Should().Be(size);
    }

    [Fact]
    public void Advanced_CorrectProperties()
    {
        var size = 64;
        var written = 14;
        var buffer = new Buffer(size, new());
        var data = Enumerable.Range(0, 64).Select(x => (byte)x).ToArray().AsSpan();

        data.CopyTo(buffer.GetSpan());
        buffer.Advance(written);

        buffer.RemainingCapacity.Should().Be(size - written);
        buffer.Written.Should().Be(written);
        buffer.GetSpan().Length.Should().Be(size - written);
        buffer.GetSpan().SequenceEqual(data.Slice(written)).Should().BeTrue();
    }

    [Fact]
    public void Flush_WrittenFlushed()
    {
        var size = 64;
        var written = 14;
        var buffer = new Buffer(size, new());
        var data = Enumerable.Range(0, 64).Select(x => (byte)x).ToArray().AsSpan();

        data.CopyTo(buffer.GetSpan());
        buffer.Advance(written);

        var target = new byte[size];
        buffer.Flush(target);

        target.Skip(written).Should().AllBeEquivalentTo(0);
        target.AsSpan(0, written).SequenceEqual(data.Slice(0, written)).Should().BeTrue();

        buffer.RemainingCapacity.Should().Be(size - written);
        buffer.Written.Should().Be(0);
        buffer.GetSpan().Length.Should().Be(size - written);
    }

    [Fact]
    public void AdvanceOffset_CorrectProperties()
    {
        var size = 64;
        var written = 14;
        var buffer = new Buffer(size, new());

        buffer.Advance(written);

        var target = new byte[written];
        buffer.Flush(target);

        var writtenMore = 5;
        buffer.Advance(writtenMore);

        buffer.RemainingCapacity.Should().Be(size - written - writtenMore);
        buffer.Written.Should().Be(writtenMore);
        buffer.GetSpan().Length.Should().Be(size - written - writtenMore);
    }

    [Fact]
    public void FlashFull_CorrectProperties()
    {
        var size = 64;
        var written = 63;
        var buffer = new Buffer(size, new());

        buffer.Advance(written);

        var target = new byte[written];
        buffer.Flush(target);

        buffer.RemainingCapacity.Should().Be(0);
        buffer.Written.Should().Be(0);
        buffer.GetSpan().Length.Should().Be(0);
    }

    [Fact]
    public async Task CompleteFlush_WrittenFlushed()
    {
        var size = 64;
        var written = 14;
        var buffer = new Buffer(size, new());
        var data = Enumerable.Range(0, 64).Select(x => (byte)x).ToArray();

        data.CopyTo(buffer.GetSpan());
        buffer.Advance(written);

        var target = new TestEntryWriter();
        await buffer.Flush(target);

        target.Items.Count.Should().Be(1);
        var flushed = target.Items[0];

        flushed.Length.Should().Be(written);
        flushed.Span.SequenceEqual(data.AsSpan(0, written)).Should().BeTrue();

        buffer.RemainingCapacity.Should().Be(0);
        buffer.Written.Should().Be(0);
        buffer.GetSpan().Length.Should().Be(0);
    }

    [Fact]
    public async Task FlushToEntryWriter_WrittenFlushed()
    {
        var size = 64;
        var written = 14;
        var buffer = new Buffer(size, new());
        var data = Enumerable.Range(0, 64).Select(x => (byte)x).ToArray();

        data.CopyTo(buffer.GetSpan());
        buffer.Advance(written);

        var target = new TestEntryWriter();
        await buffer.Flush(target);

        target.Items.Count.Should().Be(1);
        var flushed = target.Items[0];

        flushed.Length.Should().Be(written);
        flushed.Span.SequenceEqual(data.AsSpan(0, written)).Should().BeTrue();

        buffer.RemainingCapacity.Should().Be(size - written);
        buffer.Written.Should().Be(0);
        buffer.GetSpan().Length.Should().Be(size - written);
    }


    [Fact]
    public async Task DoubleFlushRenewToEntryWriter_WrittenFlushed()
    {
        var size = 64;
        var written = 14;
        var buffer = new Buffer(size, new());
        var data = Enumerable.Range(0, 64).Select(x => (byte)x).ToArray();

        data.CopyTo(buffer.GetSpan());
        buffer.Advance(written);

        var newSize = 128;
        var target = new TestEntryWriter();
        await buffer.Flush(target);

        var writtenMore = 45;
        buffer.Advance(writtenMore);
        await buffer.Flush(target);


        target.Items.Count.Should().Be(2);
        var flushed = target.Items[1];

        flushed.Length.Should().Be(writtenMore);
        flushed.Span.SequenceEqual(data.Skip(written).ToArray().AsSpan(0, writtenMore)).Should().BeTrue();

        buffer.RemainingCapacity.Should().Be(newSize);
        buffer.Written.Should().Be(0);
        buffer.GetSpan().Length.Should().Be(newSize);
    }

    [Fact]
    public async Task DoubleFlushToEntryWriter_WrittenFlushed()
    {
        var size = 64;
        var written = 14;
        var buffer = new Buffer(size, new());
        var data = Enumerable.Range(0, 64).Select(x => (byte)x).ToArray();

        data.CopyTo(buffer.GetSpan());
        buffer.Advance(written);

        var target = new TestEntryWriter();
        await buffer.Flush(target);

        var writtenMore = 5;
        buffer.Advance(writtenMore);
        await buffer.Flush(target);


        target.Items.Count.Should().Be(2);
        var flushed = target.Items[1];

        flushed.Length.Should().Be(writtenMore);
        flushed.Span.SequenceEqual(data.Skip(written).ToArray().AsSpan(0, writtenMore)).Should().BeTrue();

        buffer.RemainingCapacity.Should().Be(size - written - writtenMore);
        buffer.Written.Should().Be(0);
        buffer.GetSpan().Length.Should().Be(size - written - writtenMore);
    }

    private class TestEntryWriter : IEntryWriter
    {
        public readonly List<ReadOnlyMemory<byte>> Items = new();

        public ValueTask Write(MemoryOwner buffer)
        {
            Items.Add(buffer.Memory);
            return ValueTask.CompletedTask;
        }
    }
}
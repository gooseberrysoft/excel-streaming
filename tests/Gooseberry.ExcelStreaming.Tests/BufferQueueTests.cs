using FluentAssertions;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class BufferQueueTests
{
    [Fact]
    public void Enqueue_UsesSyncWriter_WhenTryWriteReturnsTrue()
    {
        var pool = new BufferPool(64);

        var memory = pool.Rent(16);
        var span = memory.Span;
        for (var i = 0; i < 16; i++) span[i] = (byte)i;

        var owner = new MemoryOwner(memory, 16, pool);

        var writer = new TestEntryWriter(syncAccept: true);
        var queue = new BufferQueue(writer);

        queue.Enqueue(owner);

        // Writer should have received the buffer and queue remains empty.
        writer.Items.Count.Should().Be(1);
        writer.Items[0].Length.Should().Be(16);
        queue.IsEmpty.Should().BeTrue();
    }

    [Fact]
    public void Enqueue_EnqueuesBuffer_WhenSyncWriterRejects()
    {
        var pool = new BufferPool(128);

        var memory = pool.Rent(10);
        var span = memory.Span;
        for (var i = 0; i < 10; i++) span[i] = (byte)(i + 1);

        var owner = new MemoryOwner(memory, 10, pool);

        var writer = new TestEntryWriter(syncAccept: false);
        var queue = new BufferQueue(writer);

        queue.Enqueue(owner);

        // Writer should not receive the buffer and queue should contain it.
        writer.Items.Count.Should().Be(0);
        queue.IsEmpty.Should().BeFalse();
        queue.GetLength().Should().Be(10);

        var target = new byte[10];
        queue.Flush(target, out var written);
        written.Should().Be(10);
        target.AsSpan(0, written).SequenceEqual(memory.Span.Slice(0, 10)).Should().BeTrue();
        queue.IsEmpty.Should().BeTrue();
    }

    [Fact]
    public void FlushSpan_ConcatenatesAllQueuedBuffers()
    {
        var pool = new BufferPool(64);

        var mem1 = pool.Rent(7);
        for (var i = 0; i < 7; i++) mem1.Span[i] = (byte)(i + 1);
        var owner1 = new MemoryOwner(mem1, 7, pool);

        var mem2 = pool.Rent(5);
        for (var i = 0; i < 5; i++) mem2.Span[i] = (byte)(i + 20);
        var owner2 = new MemoryOwner(mem2, 5, pool);

        var queue = new BufferQueue();
        queue.Enqueue(owner1);
        queue.Enqueue(owner2);

        var buffer = new byte[12];
        queue.Flush(buffer, out var written);

        written.Should().Be(12);
        buffer.AsSpan(0, 7).SequenceEqual(mem1.Span.Slice(0, 7)).Should().BeTrue();
        buffer.AsSpan(7, 5).SequenceEqual(mem2.Span.Slice(0, 5)).Should().BeTrue();
        queue.IsEmpty.Should().BeTrue();
    }

    [Fact]
    public async Task FlushToWriter_SingleBuffer_UsesDirectPath()
    {
        var pool = new BufferPool(32);

        var mem = pool.Rent(9);
        for (var i = 0; i < 9; i++) mem.Span[i] = (byte)(i + 2);
        var owner = new MemoryOwner(mem, 9, pool);

        var queue = new BufferQueue();
        queue.Enqueue(owner);

        var writer = new TestEntryWriter(syncAccept: false);
        await queue.Flush(writer);

        writer.Items.Count.Should().Be(1);
        writer.Items[0].Length.Should().Be(9);
        writer.Items[0].Span.SequenceEqual(mem.Span.Slice(0, 9)).Should().BeTrue();
        queue.IsEmpty.Should().BeTrue();
    }

    [Fact]
    public async Task FlushToWriter_MultipleBuffers_WritesAll()
    {
        var pool = new BufferPool(64);

        var mem1 = pool.Rent(6);
        for (var i = 0; i < 6; i++) mem1.Span[i] = (byte)(i + 10);
        var owner1 = new MemoryOwner(mem1, 6, pool);

        var mem2 = pool.Rent(4);
        for (var i = 0; i < 4; i++) mem2.Span[i] = (byte)(i + 50);
        var owner2 = new MemoryOwner(mem2, 4, pool);

        var queue = new BufferQueue();
        queue.Enqueue(owner1);
        queue.Enqueue(owner2);

        var writer = new TestEntryWriter(syncAccept: false);
        await queue.Flush(writer);

        writer.Items.Count.Should().Be(2);
        writer.Items[0].Length.Should().Be(6);
        writer.Items[1].Length.Should().Be(4);
        writer.Items[0].Span.SequenceEqual(mem1.Span.Slice(0, 6)).Should().BeTrue();
        writer.Items[1].Span.SequenceEqual(mem2.Span.Slice(0, 4)).Should().BeTrue();
        queue.IsEmpty.Should().BeTrue();
    }

    [Fact]
    public void Dispose_ReleasesQueuedBuffers()
    {
        var pool = new BufferPool(64);

        var mem1 = pool.Rent(3);
        var owner1 = new MemoryOwner(mem1, 3, pool);

        var mem2 = pool.Rent(2);
        var owner2 = new MemoryOwner(mem2, 2, pool);

        var queue = new BufferQueue();
        queue.Enqueue(owner1);
        queue.Enqueue(owner2);

        queue.IsEmpty.Should().BeFalse();
        queue.GetLength().Should().BeGreaterThan(0);

        queue.Dispose();

        queue.IsEmpty.Should().BeTrue();
        queue.GetLength().Should().Be(0);
    }

    private sealed class TestEntryWriter : IEntryWriter, ISyncEntryWriter
    {
        public readonly List<ReadOnlyMemory<byte>> Items = new();

        private readonly bool _syncAccept;

        public TestEntryWriter(bool syncAccept) => _syncAccept = syncAccept;

        public ValueTask Write(MemoryOwner buffer)
        {
            Items.Add(buffer.Memory);
            return ValueTask.CompletedTask;
        }

        public bool TryWrite(in MemoryOwner buffer)
        {
            if (_syncAccept)
                Items.Add(buffer.Memory);

            return _syncAccept;
        }
    }
}
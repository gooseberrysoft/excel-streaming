using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class BuffersChain : IDisposable
{
    private const int MinRemainingCapacity = 512;

    private readonly BufferPool _pool = new();
    private readonly Queue<MemoryOwner> _completedBuffers = new(2);
    private readonly Buffer _buffer;

    public BuffersChain(int bufferMinSize)
    {
        _buffer = new Buffer(bufferMinSize, _pool);
    }

    public int Written
    {
        get
        {
            var written = _buffer.Written;

            foreach (var buffer in _completedBuffers)
                written += buffer.Memory.Length;

            return written;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Span<byte> GetSpan(int minSize = 1)
    {
        if (_buffer.RemainingCapacity < minSize)
            _buffer.Flush(_completedBuffers, minSize);

        return _buffer.GetSpan();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Advance(int count) => _buffer.Advance(count);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ValueTask FlushCompleted(IEntryWriter output)
    {
        if (_completedBuffers.Count > 0 && _buffer.RemainingCapacity >= MinRemainingCapacity)
            return FlushCompletedBuffers(output);

        if (_completedBuffers.Count > 0)
            return FlushCompletedAsync(output);

        return FlushBuffer(output);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private ValueTask FlushBuffer(IEntryWriter output)
    {
        return _buffer.RemainingCapacity < MinRemainingCapacity
            ? _buffer.Flush(output)
            : ValueTask.CompletedTask;
    }

    private async ValueTask FlushCompletedAsync(IEntryWriter output)
    {
        await FlushCompletedBuffers(output);
        await FlushBuffer(output);
    }

    public ValueTask FlushAll(IEntryWriter output)
    {
        if (_completedBuffers.Count > 0)
            return FlushAllAsync(output);

        return _buffer.Flush(output);
    }

    private async ValueTask FlushAllAsync(IEntryWriter output)
    {
        await FlushCompletedBuffers(output);
        await _buffer.Flush(output);
    }

    public void FlushAll(Span<byte> span)
    {
        if (span.Length < Written)
            throw new ArgumentException("Span has no enough space to flush all buffers.");

        var currentPosition = 0;

        while (_completedBuffers.Count > 0)
        {
            using var buffer = _completedBuffers.Dequeue();
            var memory = buffer.Memory;

            memory.Span.CopyTo(span.Slice(currentPosition));
            currentPosition += memory.Length;
        }

        _buffer.Flush(span.Slice(currentPosition));
    }

    public void Dispose()
    {
        while (_completedBuffers.Count > 0)
            _completedBuffers.Dequeue().Dispose();

        _buffer.Dispose();
        _pool.Dispose();
    }

    private ValueTask FlushCompletedBuffers(IEntryWriter output)
    {
        if (_completedBuffers.Count == 1)
            return output.Write(_completedBuffers.Dequeue());

        return FlushCompletedBuffersAsync(output);
    }

    private async ValueTask FlushCompletedBuffersAsync(IEntryWriter output)
    {
        while (_completedBuffers.Count > 0)
            await output.Write(_completedBuffers.Dequeue());
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(ReadOnlySpan<byte> bytes)
    {
        if (!_buffer.TryWrite(bytes))
            WriteBlocks(bytes);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private void WriteBlocks(ReadOnlySpan<byte> data)
    {
        while (true)
        {
            var destination = GetSpan();
            var copyLength = Math.Min(data.Length, destination.Length);

            data.Slice(0, copyLength).CopyTo(destination);

            Advance(copyLength);

            if (data.Length == copyLength)
                return;

            data = data.Slice(copyLength);
        }
    }
}
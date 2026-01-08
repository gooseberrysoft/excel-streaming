using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class BufferSequence : IDisposable
{
    private const int MinRemainingCapacity = 128;

    private readonly BufferPool _pool = new();
    private readonly Buffer _buffer;
    private readonly BufferQueue _bufferQueue;

    public BufferSequence(int bufferMinSize) : this(bufferMinSize, new BufferQueue())
    {
    }

    public BufferSequence(int bufferMinSize, BufferQueue queue)
    {
        _buffer = new Buffer(bufferMinSize, _pool);
        _bufferQueue = queue;
    }

    public int Written => _buffer.Written + _bufferQueue.GetLength();

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Span<byte> GetSpan(int minSize = 1)
    {
        if (_buffer.RemainingCapacity < minSize)
            _buffer.Flush(_bufferQueue, minSize);

        return _buffer.GetSpan();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Advance(int count) => _buffer.Advance(count);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ValueTask FlushCompleted(IEntryWriter output)
    {
        if (!_bufferQueue.IsEmpty)
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
        await _bufferQueue.Flush(output);
        await FlushBuffer(output);
    }

    public ValueTask FlushAll(IEntryWriter output)
        => !_bufferQueue.IsEmpty ? FlushAllAsync(output) : _buffer.Flush(output);

    private async ValueTask FlushAllAsync(IEntryWriter output)
    {
        await _bufferQueue.Flush(output);
        await _buffer.Flush(output);
    }

    public void FlushAll(Span<byte> span)
    {
        if (span.Length < Written)
            throw new ArgumentException("Span has no enough space to flush all buffers.");

        _bufferQueue.Flush(span, out var written);
        _buffer.Flush(span.Slice(written));
    }

    public void Dispose()
    {
        _bufferQueue.Dispose();
        _buffer.Dispose();
        _pool.Dispose();
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
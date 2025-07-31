using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class BuffersChain : IDisposable
{
    private readonly BufferPool _pool = new();
    private readonly Queue<Buffer> _completedBuffers = new();
    private Buffer _currentBuffer;

    public BuffersChain(int initialBufferSize)
    {
        _currentBuffer = new Buffer(initialBufferSize, _pool);
    }

    public int Written
    {
        get
        {
            var written = _currentBuffer.Written;

            foreach (var buffer in _completedBuffers)
                written += buffer.Written;

            return written;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Span<byte> GetSpan(int minSize = 1)
    {
        if (_currentBuffer.RemainingCapacity < minSize)
            Allocate(minSize);

        return _currentBuffer.GetSpan();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void Allocate(int minSize)
    {
        if (!_currentBuffer.IsEmpty)
            _completedBuffers.Enqueue(_currentBuffer);

        _currentBuffer = new Buffer(minSize, _pool);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Advance(int count) => _currentBuffer.Advance(count);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ValueTask FlushCompleted(IEntryWriter output)
    {
        if (_completedBuffers.Count > 0)
            return FlushCompletedAsync(output);

        return FlushCurrent(output);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private ValueTask FlushCurrent(IEntryWriter output)
    {
        if (_currentBuffer.RemainingCapacity < 512)
            return _currentBuffer.Flush(output);

        return ValueTask.CompletedTask;
    }

    private async ValueTask FlushCompletedAsync(IEntryWriter output)
    {
        await FlushCompletedBuffers(output);
        await FlushCurrent(output);
    }

    public ValueTask FlushAll(IEntryWriter output)
    {
        if (_completedBuffers.Count > 0)
            return FlushAllAsync(output);

        return _currentBuffer.Flush(output);
    }

    private async ValueTask FlushAllAsync(IEntryWriter output)
    {
        await FlushCompletedBuffers(output);
        await FlushAll(output);
    }

    public void FlushAll(Span<byte> span)
    {
        if (span.Length < Written)
            throw new ArgumentException("Span has no enough space to flush all buffers.");

        var currentPosition = 0;

        while (_completedBuffers.Count > 0)
        {
            using var buffer = _completedBuffers.Dequeue();
            var chunk = span.Slice(currentPosition, buffer.Written);
            buffer.Flush(chunk);
            currentPosition += chunk.Length;
        }

        if (!_currentBuffer.IsEmpty)
            _currentBuffer.Flush(span.Slice(currentPosition, _currentBuffer.Written));
    }

    public void Dispose()
    {
        while (_completedBuffers.Count > 0)
            _completedBuffers.Dequeue().Dispose();

        _currentBuffer.Dispose();
        _pool.Dispose();
    }

    private async Task FlushCompletedBuffers(IEntryWriter output)
    {
        while (_completedBuffers.Count > 0)
        {
            using var buffer = _completedBuffers.Dequeue();
            await buffer.Flush(output, allocateNew: false);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(ReadOnlySpan<byte> bytes)
    {
        if (!_currentBuffer.TryWrite(bytes))
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
using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class BuffersChain : IDisposable
{
    private const int MaxBufferSize = 1024 * 1024;
    private const int FlushSize = 512 * 1024;

    private int _avgBytesPerFlush = 1024;
    private int _bytesPerFlush = 0;

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

    [MethodImpl(MethodImplOptions.NoInlining)]
    private void Allocate(int minSize)
    {
        var bufferSize = _currentBuffer.AllocatedSize;

        if (!_currentBuffer.IsEmpty)
        {
            _completedBuffers.Enqueue(_currentBuffer);
            bufferSize = GetNewSize();
        }

        _currentBuffer = new Buffer(Math.Max(bufferSize, minSize), _pool);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Advance(int count)
    {
        _currentBuffer.Advance(count);
        _bytesPerFlush += count;
    }

    public ValueTask FlushCompleted(IEntryWriter output)
    {
        _avgBytesPerFlush = (_avgBytesPerFlush + _bytesPerFlush) / 2;
        _bytesPerFlush = 0;
        
        if (_completedBuffers.Count > 0)
            return FlushCompletedAsync(output);
        
        return FlushCurrent(output);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private ValueTask FlushCurrent(IEntryWriter output)
    {
        if (_currentBuffer.RemainingCapacity <= _avgBytesPerFlush + 64)
            return _currentBuffer.CompleteFlush(output, GetNewSize());

        if (_currentBuffer.Written >= FlushSize)
            return _currentBuffer.Flush(output, GetNewSize());

        return ValueTask.CompletedTask;
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private async ValueTask FlushCompletedAsync(IEntryWriter output)
    {
        await FlushCompletedBuffers(output);
        await FlushCurrent(output);
    }

    public ValueTask FlushAll(IEntryWriter output)
    {
        if (_completedBuffers.Count > 0)
            return FlushAllAsync(output);

        return _currentBuffer.Flush(output, GetNewSize());
    }

    private int GetNewSize()
    {
        if (_currentBuffer.AllocatedSize == MaxBufferSize)
            return MaxBufferSize / 2;

        return Math.Min(MaxBufferSize, _currentBuffer.AllocatedSize * 2);
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
            await buffer.CompleteFlush(output);
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
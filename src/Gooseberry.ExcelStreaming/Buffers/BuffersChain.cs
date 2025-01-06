using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class BuffersChain : IDisposable
{
    private const int MaxBufferSize = 1024 * 1024;
    private const int RemainingFlushSize = 1 * 1024;
    private const int FlushSize = 512 * 1024;

    private int _bufferSize;

    private readonly Queue<Buffer> _completedBuffers = new();
    private Buffer _currentBuffer;

    public BuffersChain(int initialBufferSize)
    {
        _bufferSize = initialBufferSize;
        _currentBuffer = new Buffer(_bufferSize);
    }

    public int Written
    {
        get
        {
            var written = _currentBuffer.Written;

            if (_completedBuffers.Count > 0)
                written += GetCompletedBuffersWritten();

            return written;
        }
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private int GetCompletedBuffersWritten()
    {
        var written = 0;

        foreach (var buffer in _completedBuffers)
            written += buffer.Written;

        return written;
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
        if (!_currentBuffer.IsEmpty)
        {
            _completedBuffers.Enqueue(_currentBuffer);
            _bufferSize = Math.Min(MaxBufferSize, _bufferSize * 2);
        }

        _currentBuffer = new Buffer(Math.Max(_bufferSize, minSize));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Advance(int count)
        => _currentBuffer.Advance(count);

    public async ValueTask FlushCompleted(IEntryWriter output)
    {
        while (_completedBuffers.Count > 0)
        {
            using var buffer = _completedBuffers.Dequeue();
            await buffer.CompleteFlush(output);
        }

        if (_currentBuffer.RemainingCapacity <= RemainingFlushSize)
        {
            await _currentBuffer.CompleteFlush(output, newSize: Math.Min(MaxBufferSize, _bufferSize * 2));

            _bufferSize = _currentBuffer.UnderlyingLength;
        }
        else if (_currentBuffer.Written >= FlushSize)
        {
            await _currentBuffer.Flush(output, newSize: Math.Min(MaxBufferSize, _bufferSize * 2));

            _bufferSize = _currentBuffer.UnderlyingLength;
        }
    }

    public async ValueTask FlushAll(IEntryWriter output)
    {
        while (_completedBuffers.Count > 0)
        {
            using var buffer = _completedBuffers.Dequeue();
            await buffer.CompleteFlush(output);
        }

        await _currentBuffer.Flush(output, newSize: Math.Min(MaxBufferSize, _bufferSize * 2));

        _bufferSize = _currentBuffer.UnderlyingLength;
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
    }
}
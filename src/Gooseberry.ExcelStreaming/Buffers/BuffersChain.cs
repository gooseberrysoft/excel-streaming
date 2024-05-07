using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class BuffersChain : IDisposable
{
    private const int MaxBufferSize = 1024 * 1024;
    private const int MaxRemainingBytes = 4 * 1024;

    private int _bufferSize;
    private int _maxBytesPerFlush = 24;
    private int _prevWritten = 0;

    private readonly Queue<Buffer> _completedBuffers = new();
    private Buffer _currentBuffer;

    public BuffersChain(int initialBufferSize)
    {
        _bufferSize = initialBufferSize;

        _currentBuffer = new Buffer(_bufferSize);
    }

    public void SetBufferSize(int size)
        => _bufferSize = Math.Max(size, Buffer.MinSize);

    public int Written
    {
        get
        {
            var written = _currentBuffer.Written;

            if (_completedBuffers.Count > 0)
            {
                foreach (var buffer in _completedBuffers)
                    written += buffer.Written;
            }

            return written;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Span<byte> GetSpan(int minSize = 1)
    {
        if (_currentBuffer.RemainingCapacity < minSize)
        {
            if (!_currentBuffer.IsEmpty)
                _completedBuffers.Enqueue(_currentBuffer);

            _bufferSize = Math.Min(MaxBufferSize, _bufferSize * 2);
            _currentBuffer = new Buffer(Math.Max(_bufferSize, minSize));
        }

        return _currentBuffer.GetSpan();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Advance(int count)
        => _currentBuffer.Advance(count);

    public async ValueTask FlushCompleted(IEntryWriter output)
    {
        var written = Written;

        _maxBytesPerFlush = Math.Max(written - _prevWritten, (int)(_maxBytesPerFlush * 0.98));

        while (_completedBuffers.Count > 0)
        {
            using var buffer = _completedBuffers.Dequeue();
            await buffer.Flush(output, newSize: 0);
        }

        if (_currentBuffer.RemainingCapacity <= Math.Min((int)(_maxBytesPerFlush * 1.1), MaxRemainingBytes))
        {
            _bufferSize = Math.Min(MaxBufferSize, _bufferSize * 2);

            await _currentBuffer.Flush(output, newSize: _bufferSize);
        }

        _prevWritten = written;
    }

    public async ValueTask FlushAll(IEntryWriter output, int nextBufferSize)
    {
        while (_completedBuffers.Count > 0)
        {
            using var buffer = _completedBuffers.Dequeue();
            await buffer.Flush(output, newSize: 0);
        }

        _bufferSize = Math.Max(nextBufferSize, Buffer.MinSize);

        await _currentBuffer.Flush(output, newSize: nextBufferSize);
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
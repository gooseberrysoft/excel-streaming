using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class BuffersChain : IDisposable
{
    private const int MaxBufferSize = 1024 * 1024;
    private int _bufferSize;
    private readonly double _flushThreshold;

    private readonly Queue<Buffer> _completedBuffers = new(2);
    private Buffer _currentBuffer;

    public BuffersChain(int bufferSize, double flushThreshold)
    {
        if (flushThreshold is <= 0 or > 1.0)
            throw new ArgumentOutOfRangeException(nameof(flushThreshold),
                "Flush threshold should be in range (0..1].");

        _bufferSize = bufferSize;
        _flushThreshold = flushThreshold;

        var buffer = new Buffer(_bufferSize);
        _currentBuffer = buffer;
    }

    public void SetBufferSize(int size)
        => _bufferSize = size;

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
        while (_completedBuffers.Count > 0)
            await output.Write(_completedBuffers.Dequeue());

        if (_currentBuffer.Saturation >= _flushThreshold)
        {
            if (!_currentBuffer.IsEmpty)
                await output.Write(_currentBuffer);

            _bufferSize = Math.Min(MaxBufferSize, _bufferSize * 2);
            _currentBuffer = new Buffer(_bufferSize);
        }
    }

    public async ValueTask FlushAll(IEntryWriter output)
    {
        while (_completedBuffers.Count > 0)
            await output.Write(_completedBuffers.Dequeue());

        if (!_currentBuffer.IsEmpty)
            await output.Write(_currentBuffer);

        _currentBuffer = Buffer.Empty;
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
            buffer.WrittenSpan.CopyTo(chunk);
            currentPosition += chunk.Length;
        }

        if (!_currentBuffer.IsEmpty)
        {
            _currentBuffer.WrittenSpan.CopyTo(span.Slice(currentPosition, _currentBuffer.Written));
            _currentBuffer.Dispose();
            _currentBuffer = Buffer.Empty;
        }
    }

    public void Dispose()
    {
        while (_completedBuffers.Count > 0)
            _completedBuffers.Dequeue().Dispose();

        _currentBuffer.Dispose();
    }
}
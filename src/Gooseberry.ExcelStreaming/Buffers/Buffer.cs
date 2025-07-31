using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class Buffer : IDisposable
{
    private readonly BufferPool _pool;
    public const int MinSize = 32;

    private int _length;
    private Memory<byte> _buffer = default;

    public Buffer(int minSize, BufferPool pool)
    {
        _pool = pool;
        RentNew(minSize);
    }

    public bool IsEmpty => _buffer.Length == 0;

    public int RemainingCapacity => _buffer.Length - _length;

    public int Written => _length;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Span<byte> GetSpan() => _buffer.Span.Slice(_length);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Advance(int count)
    {
        if (count > RemainingCapacity)
            ThrowInvalidAdvance();

        _length += count;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool TryWrite(ReadOnlySpan<byte> bytes)
    {
        if (bytes.TryCopyTo(GetSpan()))
        {
            _length += bytes.Length;
            return true;
        }

        return false;
    }

    public ValueTask Flush(IEntryWriter output, bool allocateNew = true)
    {
        if (Written == 0)
            return ValueTask.CompletedTask;

        var memory = new MemoryOwner(_buffer, _length, _pool);
        var task = output.Write(memory);

        if (allocateNew)
            RentNew(MinSize);
        else
        {
            _length = 0;
            _buffer = default;
        }

        return task;
    }

    public void Flush(Span<byte> output)
    {
        if (Written == 0)
            return;

        _buffer.Span.Slice(0, _length).CopyTo(output);

        Return();
    }

    public void Dispose()
    {
        if (_buffer.IsEmpty)
            return;

        Return();
    }

    private void RentNew(int minSize)
    {
        _length = 0;
        _buffer = _pool.Rent(minSize);
    }

    private void Return()
    {
        if (!_buffer.IsEmpty)
            _pool.Return(_buffer);
        
        _length = 0;
        _buffer = default;
    }

    private static void ThrowInvalidAdvance()
        => throw new InvalidOperationException("Buffer haven't enough size to write data.");
}
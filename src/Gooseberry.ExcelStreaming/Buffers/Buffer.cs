using System.Buffers;
using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class Buffer : IDisposable
{
    public const int MinSize = 24;

    private readonly byte[] _buffer;
    private int _bufferIndex = 0;

    public Buffer(int size)
    {
        if (size < MinSize)
            throw new ArgumentException($"Cannot be less then {MinSize}.", nameof(size));

        _buffer = ArrayPool<byte>.Shared.Rent(size);
    }

    public int RemainingCapacity
        => _buffer.Length - _bufferIndex;

    public int Written
        => _bufferIndex;

    public double Saturation
        => (double)Written / _buffer.Length;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Span<byte> GetSpan()
        => _buffer.AsSpan(_bufferIndex);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Advance(int count)
    {
        if (count > RemainingCapacity)
            throw new InvalidOperationException($"Buffer haven't enough size to write data. " +
                $"Buffer remaining capacity is {RemainingCapacity} and data size is {count}.");

        _bufferIndex += count;
    }

    public ValueTask FlushTo(Stream stream, CancellationToken token)
    {
        var buffer = _buffer.AsMemory(0, _bufferIndex);
        _bufferIndex = 0;

        return stream.WriteAsync(buffer, token);
    }

    public void FlushTo(Span<byte> span)
    {
        _buffer.AsSpan(0, _bufferIndex).CopyTo(span);
        _bufferIndex = 0;
    }

    public void Dispose()
        => ArrayPool<byte>.Shared.Return(_buffer);
}
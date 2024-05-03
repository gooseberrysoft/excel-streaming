using System.Buffers;
using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class Buffer : IMemoryOwner<byte>
{
    public const int MinSize = 24;

    private byte[] _buffer;
    private int _bufferIndex = 0;

    public static readonly Buffer Empty = new();

    public Buffer(int size)
    {
        if (size < MinSize)
            throw new ArgumentException($"Cannot be less then {MinSize}.", nameof(size));

        _buffer = ArrayPool<byte>.Shared.Rent(size);
    }

    private Buffer()
        => _buffer = Array.Empty<byte>();

    public bool IsEmpty => _buffer.Length == 0;

    public Memory<byte> Memory => _buffer.AsMemory(0, _bufferIndex);

    public ReadOnlySpan<byte> WrittenSpan => _buffer.AsSpan(0, _bufferIndex);

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

    public void Dispose()
    {
        if (_buffer.Length == 0)
            return;

        ArrayPool<byte>.Shared.Return(_buffer);
        _buffer = Array.Empty<byte>();
        _bufferIndex = 0;
    }
}
using System.Buffers;
using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class Buffer : IDisposable
{
    public const int MinSize = 32;

    private byte[] _buffer;
    private int _bufferIndex;

    public Buffer(int size)
    {
        if (size < MinSize)
            throw new ArgumentException($"Cannot be less then {MinSize}.", nameof(size));

        _buffer = ArrayPool<byte>.Shared.Rent(size);
    }

    public bool IsEmpty => _buffer.Length == 0;

    public int RemainingCapacity
        => _buffer.Length - _bufferIndex;

    public int Written
        => _bufferIndex;

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

    public async ValueTask Flush(IEntryWriter output, int newSize)
    {
        if (_bufferIndex > 0)
            await output.Write(new MemoryOwner(_buffer, _bufferIndex));
        else if (_buffer.Length > 0)
            ArrayPool<byte>.Shared.Return(_buffer);

        _buffer = newSize == 0
            ? Array.Empty<byte>()
            : ArrayPool<byte>.Shared.Rent(newSize);

        _bufferIndex = 0;
    }

    public void Flush(Span<byte> output, int newSize = 0)
    {
        if (_bufferIndex == 0)
            return;

        _buffer.AsSpan(0, _bufferIndex).CopyTo(output);

        ArrayPool<byte>.Shared.Return(_buffer);

        _buffer = newSize == 0
            ? Array.Empty<byte>()
            : ArrayPool<byte>.Shared.Rent(newSize);

        _bufferIndex = 0;
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
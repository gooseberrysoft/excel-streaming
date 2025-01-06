using System.Buffers;
using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class Buffer : IDisposable
{
    public const int MinSize = 32;

    private int _bufferOffset;
    private int _bufferIndex;
    private byte[] _underlyingArray = [];

    private int BufferLength => _underlyingArray.Length - _bufferOffset;

    public Buffer(int size)
    {
        if (size < MinSize)
            throw new ArgumentException($"Cannot be less then {MinSize}.", nameof(size));

        RentNew(size);
    }

    public bool IsEmpty => BufferLength == 0;

    public int RemainingCapacity => BufferLength - _bufferIndex;

    public int Written => _bufferIndex;

    public int UnderlyingLength => _underlyingArray.Length;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Span<byte> GetSpan() => _underlyingArray.AsSpan(_bufferOffset + _bufferIndex);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Advance(int count)
    {
        if (count > RemainingCapacity)
            ThrowInvalidAdvance();

        _bufferIndex += count;
    }

    public async ValueTask CompleteFlush(IEntryWriter output, int newSize = 0)
    {
        if (_bufferIndex == 0)
            return;

        await output.Write(new MemoryOwner(_underlyingArray.AsMemory(_bufferOffset, _bufferIndex), _underlyingArray));

        RentNew(newSize);
    }

    public async ValueTask Flush(IEntryWriter output, int newSize)
    {
        if (_bufferIndex == 0)
            return;

        if (BufferLength - _bufferIndex <= MinSize)
        {
            await output.Write(new MemoryOwner(_underlyingArray.AsMemory(_bufferOffset, _bufferIndex), _underlyingArray));

            RentNew(newSize);
        }
        else
        {
            await output.Write(_underlyingArray.AsMemory(_bufferOffset, _bufferIndex));
            _bufferOffset += _bufferIndex;
            _bufferIndex = 0;
        }
    }

    public void Flush(Span<byte> output)
    {
        if (_bufferIndex == 0)
            return;

        _underlyingArray.AsSpan(_bufferOffset, _bufferIndex).CopyTo(output);

        _bufferOffset += _bufferIndex;
        _bufferIndex = 0;

        if (BufferLength < MinSize)
        {
            ArrayPool<byte>.Shared.Return(_underlyingArray);
            RentNew(0);
        }
    }

    public void Dispose()
    {
        if (_underlyingArray.Length == 0)
            return;

        ArrayPool<byte>.Shared.Return(_underlyingArray);

        _underlyingArray = [];
        _bufferOffset = _bufferIndex = 0;
    }

    private void RentNew(int size)
    {
        _bufferOffset = _bufferIndex = 0;
        _underlyingArray = size == 0 ? [] : ArrayPool<byte>.Shared.Rent(size);
    }

    private static void ThrowInvalidAdvance() 
        => throw new InvalidOperationException("Buffer haven't enough size to write data.");
}
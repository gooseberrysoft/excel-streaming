using System.Buffers;
using System.Runtime.CompilerServices;

// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

internal sealed class Buffer : IDisposable
{
    public const int MinSize = 32;

    private Memory<byte> _buffer;
    private int _bufferIndex;
    private byte[] _underlyingArray = Array.Empty<byte>();

    public Buffer(int size)
    {
        if (size < MinSize)
            throw new ArgumentException($"Cannot be less then {MinSize}.", nameof(size));

        RentNew(size);
    }

    public bool IsEmpty => _buffer.Length == 0;

    public int RemainingCapacity => _buffer.Length - _bufferIndex;

    public int Written => _bufferIndex;

    public int UnderlyingLength => _underlyingArray.Length;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Span<byte> GetSpan() => _buffer.Span.Slice(_bufferIndex);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Advance(int count)
    {
        if (count > RemainingCapacity)
            throw new InvalidOperationException("Buffer haven't enough size to write data.");

        _bufferIndex += count;
    }

    public async ValueTask CompleteFlush(IEntryWriter output, int newSize = 0)
    {
        if (_bufferIndex == 0)
            return;

        await output.Write(new MemoryOwner(_buffer.Slice(0, _bufferIndex), _underlyingArray));

        RentNew(newSize);
    }

    public async ValueTask Flush(IEntryWriter output, int newSize)
    {
        if (_bufferIndex == 0)
            return;

        if (_buffer.Length - _bufferIndex <= MinSize)
        {
            await output.Write(new MemoryOwner(_buffer.Slice(0, _bufferIndex), _underlyingArray));

            RentNew(newSize);
        }
        else
        {
            await output.Write(_buffer.Slice(0, _bufferIndex));
            _buffer = _buffer.Slice(_bufferIndex);
            _bufferIndex = 0;
        }
    }

    public void Flush(Span<byte> output)
    {
        if (_bufferIndex == 0)
            return;

        _buffer.Span.Slice(0, _bufferIndex).CopyTo(output);

        _buffer = _buffer.Slice(_bufferIndex);
        _bufferIndex = 0;

        if (_buffer.Length < MinSize)
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

        _buffer = _underlyingArray = Array.Empty<byte>();
        _bufferIndex = 0;
    }

    private void RentNew(int size)
    {
        _bufferIndex = 0;
        _buffer = _underlyingArray = size == 0
            ? Array.Empty<byte>()
            : ArrayPool<byte>.Shared.Rent(size);
    }
}
using System;
using System.Buffers;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Gooseberry.ExcelStreaming
{
    internal sealed class Buffer : IDisposable
    {
        private readonly byte[] _buffer;
        private int _bufferIndex = 0;

        public Buffer(int size)
        {
            if (size <= 0)
                throw new ArgumentException("Cannot be less or equal 0", nameof(size));

            _buffer = ArrayPool<byte>.Shared.Rent(size);
        }

        public int RemainingCapacity
            => _buffer.Length - _bufferIndex;

        public int Written
            => _bufferIndex;

        public double Saturation
            => (double) Written / (_buffer.Length);

        public Span<byte> GetSpan(int? sizeHint = null)
        {
            if (!sizeHint.HasValue)
                return _buffer.AsSpan(_bufferIndex);

            EnsureBufferHas(sizeHint.Value);
            return _buffer.AsSpan(_bufferIndex, sizeHint.Value);
        }

        public void Advance(int count)
        {
            EnsureBufferHas(count);
            _bufferIndex += count;
        }

        public async ValueTask FlushTo(Stream stream, CancellationToken token)
        {
            await stream.WriteAsync(_buffer.AsMemory(0, _bufferIndex), token);
            _bufferIndex = 0;
        }

        public void FlushTo(Span<byte> span)
        {
            _buffer.AsSpan(0, _bufferIndex).CopyTo(span);
            _bufferIndex = 0;
        }

        public void Dispose()
            => ArrayPool<byte>.Shared.Return(_buffer);

        private void EnsureBufferHas(int size)
        {
            if (size > RemainingCapacity)
                throw new InvalidOperationException($"Buffer haven't enough size to write data. " +
                    $"Buffer remaining capacity is {RemainingCapacity} and data size is {size}.");
        }
    }
}
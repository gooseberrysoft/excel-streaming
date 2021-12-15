using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Gooseberry.ExcelStreaming
{
    internal sealed class BuffersChain : IDisposable
    {
        private readonly int _bufferSize;
        private readonly double _flushThreshold;

        private readonly List<Buffer> _buffers = new();
        private int _currentBuffer;

        public BuffersChain(int bufferSize, double flushThreshold)
        {
            if (flushThreshold is <= 0 or > 1.0)
                throw new ArgumentOutOfRangeException(nameof(flushThreshold), "Flush threshold should be in range (0..1].");

            _bufferSize = bufferSize;
            _flushThreshold = flushThreshold;

            _buffers.Add(new Buffer(_bufferSize));
            _currentBuffer = 0;
        }

        public int Written
        {
            get
            {
                var written = 0;
                for (var i = 0; i <= _currentBuffer; i++)
                    written += _buffers[i].Written;

                return written;
            }
        }

        public Span<byte> GetSpan(int? sizeHint = null)
        {
            if (CurrentBuffer.RemainingCapacity < (sizeHint ?? 1))
                MoveToNextBuffer();

            return CurrentBuffer.GetSpan(sizeHint);
        }

        public void Advance(int count)
            => CurrentBuffer.Advance(count);

        public async ValueTask FlushCompleted(Stream stream, CancellationToken token)
        {
            for (var bufferIndex = 0; bufferIndex < _currentBuffer; bufferIndex++)
                await _buffers[bufferIndex].FlushTo(stream, token);

            (_buffers[0], _buffers[_currentBuffer]) = (_buffers[_currentBuffer], _buffers[0]);
            _currentBuffer = 0;

            if (CurrentBuffer.Saturation >= _flushThreshold)
                await CurrentBuffer.FlushTo(stream, token);
        }

        public async ValueTask FlushAll(Stream stream, CancellationToken token)
        {
            foreach (var buffer in _buffers)
                await buffer.FlushTo(stream, token);

            _currentBuffer = 0;
        }

        public void FlushAll(Span<byte> span)
        {
            if (span.Length < Written)
                throw new ArgumentException("Span has no enough space wo flush all buffers.");

            var currentPosition = 0;
            foreach (var buffer in _buffers)
            {
                var chunk = span.Slice(currentPosition, buffer.Written);
                buffer.FlushTo(chunk);
                currentPosition += chunk.Length;
            }

            _currentBuffer = 0;
        }

        public void Dispose()
        {
            foreach (var buffer in _buffers)
                buffer.Dispose();
        }

        private Buffer CurrentBuffer
            => _buffers[_currentBuffer];

        private void MoveToNextBuffer()
        {
            _currentBuffer++;
            if (_buffers.Count <= _currentBuffer)
                _buffers.Add(new Buffer(_bufferSize));
        }
    }
}
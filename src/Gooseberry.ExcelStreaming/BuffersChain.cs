using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace Gooseberry.ExcelStreaming
{
    internal sealed class BuffersChain : IDisposable
    {
        private readonly int _bufferSize;
        private readonly double _flushThreshold;

        private readonly List<Buffer> _buffers = new();
        private int _currentBufferIndex;
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
            _buffers.Add(buffer);
            _currentBufferIndex = 0;
        }

        public int Written
        {
            get
            {
                var written = 0;
                for (var i = 0; i <= _currentBufferIndex; i++)
                    written += _buffers[i].Written;

                return written;
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public Span<byte> GetSpan(int minSize)
        {
            if (_currentBuffer.RemainingCapacity < minSize)
                MoveToNextBuffer();
           
            //TODO check minsize?
            return _currentBuffer.GetSpan();
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public Span<byte> GetSpan()
        {
            if (_currentBuffer.RemainingCapacity == 0)
                MoveToNextBuffer();

            return _currentBuffer.GetSpan();
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void Advance(int count)
            => _currentBuffer.Advance(count);

        public async ValueTask FlushCompleted(Stream stream, CancellationToken token)
        {
            if (_currentBufferIndex > 0)
            {
                for (var bufferIndex = 0; bufferIndex < _currentBufferIndex; bufferIndex++)
                    await _buffers[bufferIndex].FlushTo(stream, token);

                (_buffers[0], _buffers[_currentBufferIndex]) = (_buffers[_currentBufferIndex], _buffers[0]);

                SetCurrentBuffer(0);
            }

            if (_currentBuffer.Saturation >= _flushThreshold)
                await _currentBuffer.FlushTo(stream, token);
        }

        public async ValueTask FlushAll(Stream stream, CancellationToken token)
        {
            foreach (var buffer in _buffers)
                await buffer.FlushTo(stream, token);

            SetCurrentBuffer(0);
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

            SetCurrentBuffer(0);
        }

        public void Dispose()
        {
            foreach (var buffer in _buffers)
                buffer.Dispose();
        }

        private void MoveToNextBuffer()
        {
            var newIndex = _currentBufferIndex + 1;
            if (_buffers.Count <= newIndex)
                _buffers.Add(new Buffer(_bufferSize));

            SetCurrentBuffer(newIndex);
        }

        private void SetCurrentBuffer(int newIndex)
        {
            _currentBufferIndex = newIndex;
            _currentBuffer = _buffers[newIndex];
        }
    }
}
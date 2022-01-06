using System;
using System.Buffers;
using System.Buffers.Text;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Encodings.Web;
using System.Threading;
using System.Threading.Tasks;

namespace Gooseberry.ExcelStreaming
{
    internal sealed class BufferedWriter : IDisposable
    {
        private const int StackCharsThreshold = 256;
        private static readonly int IntMaximumChars = int.MinValue.ToString().Length;
        private static readonly int LongMaximumChars = long.MinValue.ToString().Length;
        private static readonly int DecimalMaximumChars = decimal.MinValue.ToString().Length;
        private static readonly int DoubleMaximumChars = double.MinValue.ToString().Length;

        private readonly BuffersChain _buffersChain;
        private readonly Encoder _encoder = Encoding.UTF8.GetEncoder();

        public BufferedWriter(int bufferSize, double flushThreshold)
            => _buffersChain = new BuffersChain(bufferSize, flushThreshold);

        public int Written
            => _buffersChain.Written;

        public BuffersChain InternalWriter
            => _buffersChain;

        public void Write(int data)
        {
            var span = _buffersChain.GetSpan(IntMaximumChars);

            if (!Utf8Formatter.TryFormat(data, span, out var encodedBytes))
                throw new InvalidOperationException("Cannot format int to string.");

            _buffersChain.Advance(encodedBytes);
        }

        public void Write(long data)
        {
            var span = _buffersChain.GetSpan(LongMaximumChars);

            if (!Utf8Formatter.TryFormat(data, span, out var encodedBytes))
                throw new InvalidOperationException("Cannot format long to string.");

            _buffersChain.Advance(encodedBytes);
        }

        public void Write(decimal data)
        {
            var span = _buffersChain.GetSpan(DecimalMaximumChars);

            if (!Utf8Formatter.TryFormat(data, span, out var encodedBytes))
                throw new InvalidOperationException("Cannot format decimal to string.");

            _buffersChain.Advance(encodedBytes);
        }

        public void Write(DateTime data)
        {
            var span = _buffersChain.GetSpan(DoubleMaximumChars);

            if (!Utf8Formatter.TryFormat(data.ToOADate(), span, out var encodedBytes))
                throw new InvalidOperationException("Cannot format DateTime to string.");

            _buffersChain.Advance(encodedBytes);
        }

        public void Write(ReadOnlySpan<char> data)
        {
            _encoder.Reset();

            var source = data;
            var isCompleted = false;
            while (!isCompleted)
            {
                var destination = _buffersChain.GetSpan();

                _encoder.Convert(source, destination, true, out var charsConsumed, out var bytesWritten, out isCompleted);

                if (charsConsumed > 0)
                    source = source.Slice(charsConsumed);
                if (bytesWritten > 0)
                    _buffersChain.Advance(bytesWritten);
            }
        }

        [SkipLocalsInit]
        public void WriteEscaped(ReadOnlySpan<char> data)
        {
            //TODO think better about formula eg special cases when len = 1
            // we assume that 5% of symbols will be escaped by 6 characters each
            var bufferSize = Math.Min(data.Length + data.Length / 3 + 16, 32 * 1024);

            if (bufferSize <= StackCharsThreshold)
            {
                Span<char> stackBuffer = stackalloc char[bufferSize];
                WriteEscaped(data, stackBuffer);

                return;
            }

            var pooledBuffer = ArrayPool<char>.Shared.Rent(bufferSize);
            try
            {
                WriteEscaped(data, pooledBuffer);
            }
            finally
            {
                ArrayPool<char>.Shared.Return(pooledBuffer);
            }
        }

        private void WriteEscaped(ReadOnlySpan<char> data, Span<char> buffer)
        {
            var source = data;
            var lastResult = OperationStatus.DestinationTooSmall;
            while (lastResult == OperationStatus.DestinationTooSmall)
            {
                lastResult = HtmlEncoder.Default.Encode(
                    source,
                    buffer,
                    out var bytesConsumed,
                    out var bytesWritten,
                    isFinalBlock: false);

                if (lastResult == OperationStatus.InvalidData)
                    throw new InvalidOperationException($"Cannot write escaped string {data.ToString()}");

                if (bytesConsumed > 0)
                    source = source.Slice(bytesConsumed);
                if (bytesWritten > 0)
                    Write(buffer.Slice(0, bytesWritten));
            }
        }

        public void Write(ReadOnlySpan<byte> data)
        {
            var dataToWrite = data;
            while (!dataToWrite.IsEmpty)
            {
                var span = _buffersChain.GetSpan();
                var dataToIndex = Math.Min(dataToWrite.Length, span.Length);

                dataToWrite[.. dataToIndex].CopyTo(span);
                _buffersChain.Advance(dataToIndex);

                dataToWrite = dataToWrite[dataToIndex ..];
            }
        }

        public ValueTask FlushCompleted(Stream stream, CancellationToken token)
            => _buffersChain.FlushCompleted(stream, token);

        public ValueTask FlushAll(Stream stream, CancellationToken token)
            => _buffersChain.FlushAll(stream, token);

        public void FlushAll(Span<byte> span)
            => _buffersChain.FlushAll(span);

        public void Dispose()
            => _buffersChain.Dispose();
    }
}
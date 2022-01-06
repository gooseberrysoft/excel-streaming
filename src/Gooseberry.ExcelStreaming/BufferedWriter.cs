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

        public BuffersChain InternalBuffer => _buffersChain;

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

        public void WriteEscaped(ReadOnlySpan<char> data)
        {
            // we assume that 5% of symbols will be escaped by 6 characters each
            var bufferSize = Math.Min(data.Length + data.Length / 3, 32 * 1024);

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

    internal abstract class ElementWriter<T>
    {
        private readonly byte[] _prefix;
        private readonly byte[] _postfix;

        protected ElementWriter(byte[] prefix, byte[] postfix)
        {
            _prefix = prefix ?? throw new ArgumentNullException(nameof(prefix));
            _postfix = postfix;
        }

        public void Write(in T value, BuffersChain bufferWriter)
        {
            var span = bufferWriter.GetSpan();
            var written = 0;

            Write(_prefix, bufferWriter, ref span, ref written);
            WriteValue(value, bufferWriter, ref span, ref written);
            Write(_postfix, bufferWriter, ref span, ref written);

            bufferWriter.Advance(written);
        }

        protected abstract void WriteValue(
            in T value,
            BuffersChain bufferWriter,
            ref Span<byte> destination,
            ref int written);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        protected static void Write(
            ReadOnlySpan<byte> data,
            BuffersChain bufferWriter,
            ref Span<byte> destination,
            ref int written)
        {
            if (data.Length <= destination.Length)
            {
                data.CopyTo(destination);

                written += data.Length;
                destination = destination.Slice(data.Length);

                return;
            }

            WriteBlocks(data, bufferWriter, ref destination, ref written);
        }

        private static void WriteBlocks(ReadOnlySpan<byte> data, BuffersChain bufferWriter, ref Span<byte> destination, ref int written)
        {
            var remainingData = data;
            while (true)
            {
                var copyLength = Math.Min(remainingData.Length, destination.Length);

                remainingData.Slice(0, copyLength).CopyTo(destination);
                written += copyLength;

                if (remainingData.Length - copyLength == 0)
                {
                    destination = destination.Slice(copyLength);
                    return;
                }

                remainingData = remainingData.Slice(copyLength);
                bufferWriter.Advance(written);

                destination = bufferWriter.GetSpan();
                written = 0;
            }
        }
    }

    internal sealed class IntElementWriter : ElementWriter<int>
    {
        private static readonly int IntMaximumChars = int.MinValue.ToString().Length;

        public IntElementWriter(byte[] prefix, byte[] postfix) 
            : base(prefix, postfix)
        {
        }

        protected override void WriteValue(in int value, BuffersChain bufferWriter, ref Span<byte> destination, ref int written)
        {
            if (Utf8Formatter.TryFormat(value, destination, out var encodedBytes))
            {
                written += encodedBytes;
                return;
            }

            bufferWriter.Advance(written);

            destination = bufferWriter.GetSpan(IntMaximumChars);
            written = 0;

            if (!Utf8Formatter.TryFormat(value, destination, out var encodedBytes))
                throw new InvalidOperationException("Cannot format int to string.");
        }
    }
}
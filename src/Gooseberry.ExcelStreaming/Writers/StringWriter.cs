using System;
using System.Buffers;
using System.Runtime.CompilerServices;
using System.Text.Encodings.Web;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct StringWriter : IValueWriter<string>
{
    private const int StackCharsThreshold = 256;

    [SkipLocalsInit]
    public void WriteValue(in string value, BuffersChain bufferWriter, ref Span<byte> destination, ref int written)
    {
        var encoder = bufferWriter.Encoder;
        encoder.Reset();
        var data = value.AsSpan();

        //TODO think better about formula
        // we assume that 5% of symbols will be escaped by 6 characters each
        var bufferSize = Math.Min(data.Length + data.Length / 3 + 16, 32 * 1024);
        var allocOnStack = bufferSize <= StackCharsThreshold;
        var pooledBuffer = allocOnStack ? null : ArrayPool<char>.Shared.Rent(bufferSize);

        var buffer = allocOnStack ? stackalloc char[bufferSize] : pooledBuffer;

        try
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
                {
                    var sourceChars = buffer.Slice(0, bytesWritten);

                    while (true)
                    {
                        encoder.Convert(sourceChars, destination, true, out var charsConsumed,
                            out var bytesCharsWritten, out var isCompleted);
                        written += bytesCharsWritten;

                        if (isCompleted)
                        {
                            destination = destination.Slice(bytesCharsWritten);
                            break;
                        }

                        sourceChars = sourceChars.Slice(charsConsumed);
                        bufferWriter.Advance(written);

                        destination = bufferWriter.GetSpan(24);
                        written = 0;
                    }
                }
            }
        }
        finally
        {
            if (pooledBuffer != null)
                ArrayPool<char>.Shared.Return(pooledBuffer);
        }
    }
}
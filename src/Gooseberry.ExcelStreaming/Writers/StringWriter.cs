using System;
using System.Buffers;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Encodings.Web;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class StringWriter
{
    private const int StackCharsThreshold = 256;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteEscapedTo(
        this string data,
        BuffersChain buffer,
        Encoder encoder,
        ref Span<byte> destination,
        ref int written)
        => WriteEscapedTo(data.AsSpan(), buffer, encoder, ref destination, ref written);

    [SkipLocalsInit]
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteEscapedTo(
        this ReadOnlySpan<char> data,
        BuffersChain buffer,
        Encoder encoder,
        ref Span<byte> destination,
        ref int written)
    {
        encoder.Reset();

        //TODO think better about formula
        // we assume that 5% of symbols will be escaped by 6 characters each
        var bufferSize = Math.Min(data.Length + data.Length / 3 + 16, 32 * 1024);
        var allocOnStack = bufferSize <= StackCharsThreshold;
        var pooledBuffer = allocOnStack ? null : ArrayPool<char>.Shared.Rent(bufferSize);

        var charBuffer = allocOnStack ? stackalloc char[bufferSize] : pooledBuffer;

        try
        {
            var source = data;
            var lastResult = OperationStatus.DestinationTooSmall;
            while (lastResult == OperationStatus.DestinationTooSmall)
            {
                lastResult = HtmlEncoder.Default.Encode(
                    source,
                    charBuffer,
                    out var bytesConsumed,
                    out var bytesWritten,
                    isFinalBlock: false);

                if (lastResult == OperationStatus.InvalidData)
                    throw new InvalidOperationException($"Cannot write escaped string {data.ToString()}");

                if (bytesConsumed > 0)
                    source = source.Slice(bytesConsumed);
                if (bytesWritten > 0)
                {
                    var sourceChars = charBuffer.Slice(0, bytesWritten);

                    while (true)
                    {
                        encoder.Convert(
                            sourceChars, 
                            destination, 
                            flush: true, 
                            out var charsConsumed,
                            out var bytesCharsWritten, 
                            out var isCompleted);
                        written += bytesCharsWritten;

                        if (isCompleted)
                        {
                            destination = destination.Slice(bytesCharsWritten);
                            break;
                        }

                        sourceChars = sourceChars.Slice(charsConsumed);
                        buffer.Advance(written);

                        destination = buffer.GetSpan(24);
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
    
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this string data,
        BuffersChain buffer,
        Encoder encoder,
        ref Span<byte> destination,
        ref int written)
        => WriteTo(data.AsSpan(), buffer, encoder, ref destination, ref written);
    
    [SkipLocalsInit]
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this ReadOnlySpan<char> value, 
        BuffersChain bufferWriter, 
        Encoder encoder, 
        ref Span<byte> destination,
        ref int written)
    {
        encoder.Reset();

        var sourceChars = value;

        while (true)
        {
            encoder.Convert(
                sourceChars, 
                destination, 
                flush: true, 
                out var charsConsumed,
                out var bytesCharsWritten, 
                out var isCompleted);
            
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
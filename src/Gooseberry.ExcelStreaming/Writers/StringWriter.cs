using System.Buffers;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Encodings.Web;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class StringWriter
{
    private const int StackBytesThreshold = 512;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteEscapedTo(
        this string data,
        BuffersChain buffer,
        Encoder encoder,
        ref Span<byte> destination,
        ref int written)
        => WriteEscapedTo(data.AsSpan(), buffer, encoder, ref destination, ref written);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteEscapedTo(
        this ReadOnlySpan<char> data,
        BuffersChain buffer,
        Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        WriteEscapedTo(data, buffer, encoder, ref span, ref written);

        buffer.Advance(written);
    }

    internal static void WriteEscapedTo(
        this ReadOnlySpan<char> data,
        BuffersChain buffer,
        Encoder encoder,
        ref Span<byte> destination,
        ref int written)
    {
        var multiByteChar = false;

        while (true)
        {
            if (destination.Length == 0 || multiByteChar)
            {
                buffer.Advance(written);

                destination = buffer.GetSpan(Buffer.MinSize);
                written = 0;
            }

            encoder.Convert(
                data,
                destination,
                flush: true,
                out var charsConsumed,
                out var bytesWritten,
                out var isCompleted);

            var indexToEncode = HtmlEncoder.Default.FindFirstCharacterToEncodeUtf8(destination[..bytesWritten]);

            if (indexToEncode == -1)
            {
                written += bytesWritten;
                destination = destination[bytesWritten..];
            }
            else
            {
                var bytesUtf8ToEncode = destination[indexToEncode..bytesWritten];

                written += indexToEncode;
                destination = destination[indexToEncode..];

                EscapeUtf8(bytesUtf8ToEncode, buffer, ref destination, ref written);
            }

            if (isCompleted)
                break;

            data = data.Slice(charsConsumed);
            multiByteChar = destination.Length is > 0 and < 3 && data.Length > 0;
        }
    }

    [SkipLocalsInit]
    private static void EscapeUtf8(ReadOnlySpan<byte> bytesUtf8ToEncode, BuffersChain buffer, ref Span<byte> destination, ref int written)
    {
        byte[]? utf8BytesPooled = null;
        var length = bytesUtf8ToEncode.Length;

        var utf8Bytes = length <= StackBytesThreshold
            ? stackalloc byte[length]
            : (utf8BytesPooled = ArrayPool<byte>.Shared.Rent(length)).AsSpan(0, length);

        //because destination and bytesUtf8ToEncode are the same memory 
        bytesUtf8ToEncode.CopyTo(utf8Bytes);

        try
        {
            WriteEscapedUtf8To(utf8Bytes, buffer, ref destination, ref written);
        }
        finally
        {
            if (utf8BytesPooled != null)
                ArrayPool<byte>.Shared.Return(utf8BytesPooled);
        }
    }

    [SkipLocalsInit]
    internal static void WriteEscapedUtf8To(
        this scoped ReadOnlySpan<byte> utf8Data,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
    {
        var lastResult = OperationStatus.DestinationTooSmall;

        while (lastResult == OperationStatus.DestinationTooSmall)
        {
            lastResult = HtmlEncoder.Default.EncodeUtf8(
                utf8Data,
                destination,
                out var bytesConsumed,
                out var bytesWritten,
                isFinalBlock: false);

            if (lastResult == OperationStatus.InvalidData)
                throw new InvalidOperationException($"Cannot write escaped string {Encoding.UTF8.GetString(utf8Data)}");

            written += bytesWritten;
            destination = destination[bytesWritten..];

            if (lastResult == OperationStatus.Done)
                return;

            if (bytesConsumed > 0)
                utf8Data = utf8Data.Slice(bytesConsumed);

            buffer.Advance(written);

            destination = buffer.GetSpan(Buffer.MinSize);
            written = 0;
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

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this ReadOnlySpan<char> data,
        BuffersChain buffer,
        Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        WriteTo(data, buffer, encoder, ref span, ref written);

        buffer.Advance(written);
    }

    internal static void WriteTo(
        this ReadOnlySpan<char> value,
        BuffersChain buffer,
        Encoder encoder,
        ref Span<byte> destination,
        ref int written)
    {
        var sourceChars = value;
        var multiByteChar = false;

        while (true)
        {
            if (destination.Length == 0 || multiByteChar)
            {
                buffer.Advance(written);

                destination = buffer.GetSpan(Buffer.MinSize);
                written = 0;
            }

            encoder.Convert(
                sourceChars,
                destination,
                flush: true,
                out var charsConsumed,
                out var bytesCharsWritten,
                out var isCompleted);

            written += bytesCharsWritten;
            destination = destination.Slice(bytesCharsWritten);

            if (isCompleted)
                break;

            sourceChars = sourceChars.Slice(charsConsumed);

            multiByteChar = destination.Length is > 0 and < 3 && sourceChars.Length > 0;
        }
    }
}
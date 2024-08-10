using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct NumberWriter<T, TFormatter>
    where TFormatter : INumberFormatter<T>, new()
{
    private readonly TFormatter _formatter;

    public NumberWriter()
        => _formatter = new();

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteValue(in T value, BuffersChain bufferWriter, ref Span<byte> destination, ref int written)
    {
        if (_formatter.TryFormat(value, destination, out var encodedBytes))
        {
            destination = destination.Slice(encodedBytes);
            written += encodedBytes;
            return;
        }

        WriteAdvance(value, bufferWriter, ref destination, ref written);
    }

    private void WriteAdvance(in T value, BuffersChain bufferWriter, ref Span<byte> destination, ref int written)
    {
        bufferWriter.Advance(written);
        destination = bufferWriter.GetSpan(_formatter.MaximumChars);
        written = 0;

        var encodedBytes = 0;
        var attempt = 1;

        while (!_formatter.TryFormat(value, destination, out encodedBytes))
        {
            if (attempt > 10)
                throw new InvalidOperationException($"Can't format {typeof(T)}. Not enough memory");

            attempt++;
            destination = bufferWriter.GetSpan(_formatter.MaximumChars * attempt);
        }

        destination = destination.Slice(encodedBytes);
        written += encodedBytes;
    }
}
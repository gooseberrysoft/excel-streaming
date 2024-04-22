namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct NumberWriter<T, TFormatter>
    where TFormatter : INumberFormatter<T>, new()
{
    private readonly TFormatter _formatter;

    public NumberWriter() 
        => _formatter = new();

    public void WriteValue(in T value, BuffersChain bufferWriter, ref Span<byte> destination, ref int written)
    {
        if (_formatter.TryFormat(value, destination, out var encodedBytes))
        {
            destination = destination.Slice(encodedBytes);
            written += encodedBytes;
            return;
        }

        bufferWriter.Advance(written);
        destination = bufferWriter.GetSpan(_formatter.MaximumChars);
        written = 0;

        if (!_formatter.TryFormat(value, destination, out encodedBytes))
            throw new InvalidOperationException($"Can't format {typeof(T)}. Not enough memory");

        destination = destination.Slice(encodedBytes);
        written += encodedBytes;
    }
}
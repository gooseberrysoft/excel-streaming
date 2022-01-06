using System;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class ElementWriter<T, TWriter>
    where TWriter : IValueWriter<T>, new()
{
    private readonly byte[] _prefix;
    private readonly byte[] _postfix;
    private readonly TWriter _valueWriter;

    public ElementWriter(byte[] prefix, byte[] postfix)
    {
        _prefix = prefix ?? throw new ArgumentNullException(nameof(prefix));
        _postfix = postfix;
        _valueWriter = new();
    }

    public void Write(in T value, BuffersChain bufferWriter)
    {
        var span = bufferWriter.GetSpan();
        var written = 0;

        _prefix.WriteTo( bufferWriter, ref span, ref written);
        _valueWriter.WriteValue(value, bufferWriter, ref span, ref written);
        _postfix.WriteTo(bufferWriter, ref span, ref written);

        bufferWriter.Advance(written);
    }
}
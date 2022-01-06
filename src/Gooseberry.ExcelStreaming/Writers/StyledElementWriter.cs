//using System;
//using Gooseberry.ExcelStreaming.Styles;

//namespace Gooseberry.ExcelStreaming.Writers;

//internal sealed class StyledElementWriter<T, TWriter>
//    where TWriter : IValueWriter<T>, new()
//{
//    private readonly byte[] _prefix;
//    private readonly byte[] _stylePostfix;
//    private readonly byte[] _postfix;
//    private readonly TWriter _valueWriter;
//    private static readonly IntFormatter StyleFormatter = new();

//    public StyledElementWriter(byte[] prefix, byte[] stylePostfix, byte[] postfix)
//    {
//        _prefix = prefix ?? throw new ArgumentNullException(nameof(prefix));
//        _stylePostfix = stylePostfix;
//        _postfix = postfix;
//        _valueWriter = new();
//    }

//    public void Write(in T value, StyleReference style, BuffersChain bufferWriter)
//    {
//        var span = bufferWriter.GetSpan();
//        var written = 0;

//        _prefix.WriteTo(bufferWriter, ref span, ref written);
//        StyleFormatter.WriteValue(style.Value, bufferWriter, ref span, ref written);
//        _stylePostfix.WriteTo(bufferWriter, ref span, ref written);
//        _valueWriter.WriteValue(value, bufferWriter, ref span, ref written);
//        _postfix.WriteTo(bufferWriter, ref span, ref written);

//        bufferWriter.Advance(written);
//    }
//}
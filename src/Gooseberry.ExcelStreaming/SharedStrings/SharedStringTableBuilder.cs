using System.Collections.Generic;
using System.Text;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.SharedStrings;

public sealed class SharedStringTableBuilder
{
    private readonly List<string> _strings = new();

    internal static readonly SharedStringTable Default = new SharedStringTableBuilder().Build();

    public SharedStringReference GetOrAdd(string value)
    {
        var index = _strings.IndexOf(value);
        if (index >= 0)
            return new SharedStringReference(index);
        
        _strings.Add(value);
        return new SharedStringReference(_strings.Count - 1);
    }

    public SharedStringTable Build()
    {
        using var buffer = new BuffersChain(bufferSize: 4 * 1024, flushThreshold: 1.0);

        WriteSharedStrings(buffer, Encoding.UTF8.GetEncoder());
        
        var preparedData = new byte[buffer.Written];
        buffer.FlushAll(preparedData);
        return new SharedStringTable(preparedData);
    }

    private void WriteSharedStrings(BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;
        
        Constants.SharedStringTable.Prefix.WriteTo(buffer, ref span, ref written);

        foreach (var value in _strings)
        {
            Constants.SharedStringTable.Item.Prefix.WriteTo(buffer, ref span, ref written);
            value.WriteEscapedTo(buffer, encoder, ref span, ref written);
            Constants.SharedStringTable.Item.Postfix.WriteTo(buffer, ref span, ref written);
        }
        
        Constants.SharedStringTable.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
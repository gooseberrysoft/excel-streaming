using System.Collections.Generic;
using System.Text;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct SharedStringWriter
{
    public void WritePrefix(BuffersChain buffer)
        => Constants.SharedStringTable.Prefix.WriteTo(buffer);

    public void Write(string value, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;
        
        Constants.SharedStringTable.Item.Prefix.WriteTo(buffer, ref span, ref written);
        value.WriteEscapedTo(buffer, encoder, ref span, ref written);
        Constants.SharedStringTable.Item.Postfix.WriteTo(buffer, ref span, ref written);
        
        buffer.Advance(written);
    }
    
    public void WritePostfix(BuffersChain buffer)
        => Constants.SharedStringTable.Postfix.WriteTo(buffer);
}
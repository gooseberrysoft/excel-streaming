using System.Linq;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class NumberCellWriter<T, TFormatter>
    where TFormatter : INumberFormatter<T>, new()
{
    private readonly TFormatter _formatter = new();
    
    private readonly byte[] _statelessPrefix;
    private readonly byte[] _statePrefix;
    private readonly byte[] _statePostfix;
    
    public NumberCellWriter(byte[] dataType)
    {
        _statelessPrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
            .Concat(dataType)
            .Concat(Constants.Worksheet.SheetData.Row.Cell.Middle)
            .ToArray();

        _statePrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
            .Concat(dataType)
            .Concat(Constants.Worksheet.SheetData.Row.Cell.Style.Prefix)
            .ToArray();
        
        _statePostfix = Constants.Worksheet.SheetData.Row.Cell.Style.Postfix
            .Concat(Constants.Worksheet.SheetData.Row.Cell.Middle)
            .ToArray();
    }
    
    public void Write(in T value, BuffersChain buffer, StyleReference? style = null)
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            _statePrefix.WriteTo(buffer, ref span, ref written);
            style.Value.Value.WriteTo(buffer, ref span, ref written);
            _statePostfix.WriteTo(buffer, ref span, ref written);
            value.WriteTo(_formatter, buffer, ref span, ref written);
            Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);
            
            buffer.Advance(written);
            return;
        }

        _statelessPrefix.WriteTo(buffer, ref span, ref written);
        value.WriteTo(_formatter, buffer, ref span, ref written);
        Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);
            
        buffer.Advance(written);
    }
}
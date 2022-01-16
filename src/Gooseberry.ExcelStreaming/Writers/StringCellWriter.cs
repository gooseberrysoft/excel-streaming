using System;
using System.Linq;
using System.Text;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class StringCellWriter
{
    private readonly EscapedStringWriter _valueWriter = new();
    private readonly NumberWriter<int, IntFormatter> _styleWriter = new();

    private readonly byte[] _statelessPrefix;
    private readonly byte[] _statePrefix;
    private readonly byte[] _statePostfix;

    public StringCellWriter()
    {
        _statelessPrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
            .Concat(Constants.Worksheet.SheetData.Row.Cell.StringDataType)
            .Concat(Constants.Worksheet.SheetData.Row.Cell.Middle)
            .ToArray();

        _statePrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
            .Concat(Constants.Worksheet.SheetData.Row.Cell.StringDataType)
            .Concat(Constants.Worksheet.SheetData.Row.Cell.Style.Prefix)
            .ToArray();
        
        _statePostfix = Constants.Worksheet.SheetData.Row.Cell.Style.Postfix
            .Concat(Constants.Worksheet.SheetData.Row.Cell.Middle)
            .ToArray();
    }
    
    public void Write(ReadOnlySpan<char> value, BuffersChain bufferWriter, Encoder encoder, StyleReference? style = null)
    {
        var span = bufferWriter.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            _statePrefix.WriteTo(bufferWriter, ref span, ref written);
            _styleWriter.WriteValue(style.Value.Value, bufferWriter, ref span, ref written);
            _statePostfix.WriteTo(bufferWriter, ref span, ref written);
            _valueWriter.WriteValue(value, bufferWriter, encoder, ref span, ref written);
            Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(bufferWriter, ref span, ref written);
            
            bufferWriter.Advance(written);
            return;
        }

        _statelessPrefix.WriteTo(bufferWriter, ref span, ref written);
        _valueWriter.WriteValue(value, bufferWriter, encoder, ref span, ref written);
        Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(bufferWriter, ref span, ref written);
            
        bufferWriter.Advance(written);        
    }
}
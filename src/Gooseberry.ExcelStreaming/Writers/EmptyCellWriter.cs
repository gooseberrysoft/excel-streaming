using System.Linq;
using Gooseberry.ExcelStreaming.Styles;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class EmptyCellWriter
{
    private readonly byte[] _stateless;
    private readonly byte[] _statePrefix;
    private readonly byte[] _statePostfix;

    public EmptyCellWriter()
    {
        _stateless = Constants.Worksheet.SheetData.Row.Cell.Prefix
            .Concat(Constants.Worksheet.SheetData.Row.Cell.StringDataType)
            .Concat(Constants.Worksheet.SheetData.Row.Cell.Middle)
            .Concat(Constants.Worksheet.SheetData.Row.Cell.Postfix)
            .ToArray();

        _statePrefix = Constants.Worksheet.SheetData.Row.Cell.Prefix
            .Concat(Constants.Worksheet.SheetData.Row.Cell.StringDataType)
            .Concat(Constants.Worksheet.SheetData.Row.Cell.Style.Prefix)
            .ToArray();
        
        _statePostfix = Constants.Worksheet.SheetData.Row.Cell.Style.Postfix
            .Concat(Constants.Worksheet.SheetData.Row.Cell.Middle)
            .ToArray();
    }

    public void Write(BuffersChain buffer, StyleReference? style = null)
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (style.HasValue)
        {
            _statePrefix.WriteTo(buffer, ref span, ref written);
            style.Value.Value.WriteTo(buffer, ref span, ref written);
            _statePostfix.WriteTo(buffer, ref span, ref written);
            Constants.Worksheet.SheetData.Row.Cell.Postfix.WriteTo(buffer, ref span, ref written);
            
            buffer.Advance(written);
            return;
        }

        _stateless.WriteTo(buffer, ref span, ref written);
        buffer.Advance(written);         
    }
}
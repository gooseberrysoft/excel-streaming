using Gooseberry.ExcelStreaming.Pictures.Abstractions;

namespace Gooseberry.ExcelStreaming.Writers;

internal sealed class AnchorCellWriter
{
    public void Write(in AnchorCell cell, BuffersChain buffer)
    {
        var written = 0;
        var span = buffer.GetSpan();

        Write(cell, buffer, ref span, ref written);
        buffer.Advance(written);
    }

    public void Write(in AnchorCell cell, BuffersChain buffer, ref Span<byte> span, ref int written)
    {
        Constants.Drawing.AnchorCell.Column.GetPrefix().WriteTo(buffer, ref span, ref written);
        cell.Column.WriteTo(buffer, ref span, ref written);
        Constants.Drawing.AnchorCell.Column.GetPostfix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.AnchorCell.Row.GetPrefix().WriteTo(buffer, ref span, ref written);
        cell.Row.WriteTo(buffer, ref span, ref written);
        Constants.Drawing.AnchorCell.Row.GetPostfix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.AnchorCell.ColumnOffset.GetPrefix().WriteTo(buffer, ref span, ref written);
        cell.Offset.X.WriteTo(buffer, ref span, ref written);
        Constants.Drawing.AnchorCell.ColumnOffset.GetPostfix().WriteTo(buffer, ref span, ref written);

        Constants.Drawing.AnchorCell.RowOffset.GetPrefix().WriteTo(buffer, ref span, ref written);
        cell.Offset.Y.WriteTo(buffer, ref span, ref written);
        Constants.Drawing.AnchorCell.RowOffset.GetPostfix().WriteTo(buffer, ref span, ref written);
    }
}

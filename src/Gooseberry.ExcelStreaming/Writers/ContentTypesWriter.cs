using System.Text;
using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Pictures;

namespace Gooseberry.ExcelStreaming.Writers;

internal readonly struct ContentTypesWriter
{
    public void Write(IReadOnlyCollection<Sheet> sheets, SheetDrawings sheetDrawings, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Constants.XmlPrefix.WriteTo(buffer, ref span, ref written);
        Constants.ContentTypes.Prefix.WriteTo(buffer, ref span, ref written);

        foreach (var pictureFormat in sheetDrawings.Pictures.Select(v => v.Format).Distinct())
        {
            Constants.ContentTypes.Get(pictureFormat).WriteTo(buffer, ref span, ref written);
        }

        foreach (var sheet in sheets)
        {
            Constants.ContentTypes.Sheet.Prefix.WriteTo(buffer, ref span, ref written);
            sheet.RelationshipId.WriteTo(buffer, encoder, ref span, ref written);
            Constants.ContentTypes.Sheet.Postfix.WriteTo(buffer, ref span, ref written);

            var drawing = sheetDrawings.Get(sheet.Id);

            if (!drawing.IsEmpty)
            {
                Constants.ContentTypes.Drawing.GetPrefix().WriteTo(buffer, ref span, ref written);
                PathResolver.GetDrawingFullPath(drawing).EnsureLeadingSlash().WriteTo(buffer, encoder, ref span, ref written);
                Constants.ContentTypes.Drawing.GetPostfix().WriteTo(buffer, ref span, ref written);
            }
        }

        Constants.ContentTypes.Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}

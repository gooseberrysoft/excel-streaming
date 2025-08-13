using System.Text;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class WorkbookWriter
{
    private static ReadOnlySpan<byte> Prefix
        => "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheets>"u8;

    private static ReadOnlySpan<byte> Postfix => "</sheets></workbook>"u8;

    private static ReadOnlySpan<byte> SheetStartPrefix => "<sheet name=\""u8;

    private static ReadOnlySpan<byte> SheetEndPrefix => "\" sheetId=\""u8;

    private static ReadOnlySpan<byte> SheetEndPostfix => "\" r:id=\"sheet"u8;

    private static ReadOnlySpan<byte> SheetPostfix =>
        "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>"u8;


    public static void Write(IReadOnlyCollection<Sheet> sheets, BuffersChain buffer, Encoder encoder)
    {
        var span = buffer.GetSpan();
        var written = 0;

        Prefix.WriteTo(buffer, ref span, ref written);

        foreach (var sheet in sheets)
        {
            SheetStartPrefix.WriteTo(buffer, ref span, ref written);
            sheet.Name.WriteEscapedTo(buffer, encoder, ref span, ref written);
            SheetEndPrefix.WriteTo(buffer, ref span, ref written);

            sheet.Id.WriteTo(buffer, ref span, ref written);
            SheetEndPostfix.WriteTo(buffer, ref span, ref written);
            sheet.Id.WriteTo(buffer, ref span, ref written);
            SheetPostfix.WriteTo(buffer, ref span, ref written);
        }

        Postfix.WriteTo(buffer, ref span, ref written);

        buffer.Advance(written);
    }
}
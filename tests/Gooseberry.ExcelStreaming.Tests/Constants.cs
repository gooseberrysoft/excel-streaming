using Gooseberry.ExcelStreaming.Styles;
using Color = System.Drawing.Color;

namespace Gooseberry.ExcelStreaming.Tests;

public static class Constants
{
    public static Style DefaultNumberStyle = new(
        Format: "General",
        Borders: new Borders(),
        Font: new Font(Size: 11, Name: null, Color: Color.Black, Bold: false),
        Fill: new Fill(Pattern: FillPattern.None, Color: null));

    public static Style DefaultDateTimeStyle = new(
        Format: "dd.mm.yyyy",
        Borders: new Borders(),
        Font: new Font(Size: 11, Name: null, Color: Color.Black, Bold: false),
        Fill: new Fill(Pattern: FillPattern.None, Color: null));

    public static Style DefaultHyperlinkStyle = new(
        Format: "General",
        Borders: new Borders(),
        Font: new Font(Size: 11, Name: null, Color: Color.Navy, Bold: false, Underline: Underline.Single),
        Fill: new Fill(Pattern: FillPattern.None, Color: null));
}
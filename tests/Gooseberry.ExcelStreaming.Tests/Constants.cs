using Gooseberry.ExcelStreaming.Styles;
using Color = System.Drawing.Color;

namespace Gooseberry.ExcelStreaming.Tests;

public static class Constants
{
    public static Style DefaultNumberStyle = new (
        format: "General",
        borders: new Borders(),
        font: new Font(size: 11, name: null, color: Color.Black, bold: false),
        fill: new Fill(pattern: FillPattern.None, color: null));

    public static Style DefaultDateTimeStyle = new (
        format: "dd.mm.yyyy",
        borders: new Borders(),
        font: new Font(size: 11, name: null, color: Color.Black, bold: false),
        fill: new Fill(pattern: FillPattern.None, color: null));
        
    public static Style DefaultHyperlinkStyle = new (
        format: "General",
        borders: new Borders(),
        font: new Font(size: 11, name: null, color: 0x0000007b, bold: false, underline: Underline.Single),
        fill: new Fill(pattern: FillPattern.None, color: null));
}
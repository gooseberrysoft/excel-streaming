using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles;

[StructLayout(LayoutKind.Auto)]
public readonly struct Style
{
    public Style(
        Font? font = null,
        Fill? fill = null,
        Borders? borders = null,
        string? format = null,
        Alignment? alignment = null)
    {
            Alignment = alignment;
            Font = font;
            Fill = fill;
            Borders = borders;
            Format = format;
        }

    public Font? Font { get; }

    public Fill? Fill { get; }

    public Borders? Borders { get; }

    public string? Format { get; }

    public Alignment? Alignment { get; }
}
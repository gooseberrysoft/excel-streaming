using System.Drawing;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles;

[StructLayout(LayoutKind.Auto)]
public readonly record struct Border(BorderStyle Style, Color Color)
{
    public Border() : this(BorderStyle.Thin)
    {
    }

    public Border(Color color) : this(BorderStyle.Thin, color)
    {
    }

    public Border(BorderStyle style) : this(style, Color.Black)
    {
    }
};
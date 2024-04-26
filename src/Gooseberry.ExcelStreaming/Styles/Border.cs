using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles;

[StructLayout(LayoutKind.Auto)]
public readonly struct Border : IEquatable<Border>
{
    public Border(BorderStyle style, Color color)
    {
            Style = style;
            Color = color;
        }

    public BorderStyle Style { get; }

    public Color Color { get; }

    public bool Equals(Border other)
        => Style == other.Style && Color.Equals(other.Color);

    public override bool Equals(object? other)
        => other is Border border && Equals(border);

    public override int GetHashCode()
        => HashCode.Combine(Style, Color);
}
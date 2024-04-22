using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct Alignment : IEquatable<Alignment>
    {
        public Alignment(HorizontalAlignment? horizontal, VerticalAlignment? vertical, bool wrapText)
        {
            Horizontal = horizontal;
            Vertical = vertical;
            WrapText = wrapText;
        }

        public HorizontalAlignment? Horizontal { get; }

        public VerticalAlignment? Vertical { get; }

        public bool WrapText { get; }

        public bool Equals(Alignment other)
        {
            return Nullable.Equals(Horizontal, other.Horizontal) &&
               Nullable.Equals(Vertical, other.Vertical) &&
               WrapText == other.WrapText;
        }

        public override bool Equals(object? other)
            => other is Alignment alignment && Equals(alignment);

        public override int GetHashCode()
            => HashCode.Combine(Horizontal, Vertical, WrapText);
    }
}

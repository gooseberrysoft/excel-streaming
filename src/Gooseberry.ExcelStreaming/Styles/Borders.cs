using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct Borders : IEquatable<Borders>
    {
        public Borders(Border? left, Border? right, Border? top, Border? bottom)
        {
            Left = left;
            Right = right;
            Top = top;
            Bottom = bottom;
        }

        public Border? Left { get; }

        public Border? Right { get; }

        public Border? Top { get; }

        public Border? Bottom { get; }

        public bool Equals(Borders other)
        {
            return Nullable.Equals(Left, other.Left) &&
               Nullable.Equals(Right, other.Right) &&
               Nullable.Equals(Top, other.Top) &&
               Nullable.Equals(Bottom, other.Bottom);
        }

        public override bool Equals(object? other)
            => other is Borders borders && Equals(borders);

        public override int GetHashCode()
            => HashCode.Combine(Left, Right, Top, Bottom);
    }
}

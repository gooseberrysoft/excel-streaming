using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct Fill : IEquatable<Fill>
    {
        public Fill(Color? color = null, FillPattern pattern = FillPattern.Solid)
        {
            Color = color;
            Pattern = pattern;
        }

        public Color? Color { get; }

        public FillPattern Pattern { get; }

        public bool Equals(Fill other)
            => Nullable.Equals(Color, other.Color) && Pattern == other.Pattern;

        public override bool Equals(object? other)
            => other is Fill fill && Equals(fill);

        public override int GetHashCode()
            => HashCode.Combine(Color, Pattern);
    }
}

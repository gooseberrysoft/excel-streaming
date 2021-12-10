using System;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct Font: IEquatable<Font>
    {
        public Font(int size, string? name, Color? color, bool bold)
        {
            Size = size;
            Name = name;
            Color = color;
            Bold = bold;
        }

        public int Size { get; }

        public string? Name { get; }

        public Color? Color { get; }

        public bool Bold { get; }

        public bool Equals(Font other)
        {
            return Size == other.Size &&
               string.Equals(Name, other.Name, StringComparison.Ordinal) &&
               Nullable.Equals(Color, other.Color) &&
               Bold == other.Bold;
        }

        public override bool Equals(object? other)
            => other is Font font && Equals(font);

        public override int GetHashCode()
            => HashCode.Combine(Size, Name, Color, Bold);
    }
}

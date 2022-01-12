using System;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct Font: IEquatable<Font>
    {
        public Font(
            int size = 11, 
            string? name = null, 
            Color? color = null, 
            bool bold = false,
            bool italic = false,
            bool strike = false,
            Underline underline = Underline.None)
        {
            Size = size;
            Name = name;
            Color = color;
            Bold = bold;
            Italic = italic;
            Strike = strike;
            Underline = underline;
        }

        public int Size { get; }

        public string? Name { get; }

        public Color? Color { get; }

        public bool Bold { get; }
        
        public bool Italic { get; }
        
        public bool Strike { get; }
        
        public Underline Underline { get; }

        public bool Equals(Font other)
        {
            return Size == other.Size &&
               string.Equals(Name, other.Name, StringComparison.Ordinal) &&
               Nullable.Equals(Color, other.Color) &&
               Bold == other.Bold &&
               Italic == other.Italic &&
               Strike == other.Strike &&
               Underline == other.Underline;
        }

        public override bool Equals(object? other)
            => other is Font font && Equals(font);

        public override int GetHashCode()
            => HashCode.Combine(Size, Name, Color, Bold, Italic, Strike, Underline);
    }
}

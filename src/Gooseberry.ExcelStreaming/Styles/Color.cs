using System;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Styles
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct Color : IEquatable<Color>
    {
        private readonly uint _argb;

        public Color(uint argb)
            => _argb = argb;

        public override string ToString()
            => _argb.ToString("x8");

        public static implicit operator Color(uint argb)
            => new (argb);

        public static implicit operator Color(System.Drawing.Color color)
            => new((uint)color.ToArgb());

        public bool Equals(Color other)
            => _argb == other._argb;

        public override bool Equals(object? other)
            => other is Color color && Equals(color);

        public override int GetHashCode()
            => (int)_argb;
    }
}

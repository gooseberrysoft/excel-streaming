using System.Text;

namespace Gooseberry.ExcelStreaming.Styles;

public readonly struct BorderStyle : IEquatable<BorderStyle>
{
    private readonly byte[] _value;

    internal BorderStyle(byte[] value)
        => _value = value;

    internal byte[] Value => _value ?? None.Value;

    public override string ToString()
        => Encoding.UTF8.GetString(Value);

    public bool Equals(BorderStyle other)
        => Value.AsSpan().SequenceEqual(other.Value);

    public override bool Equals(object? obj)
        => obj is BorderStyle other && Equals(other);

    public override int GetHashCode()
    {
        var hashCode = new HashCode();
        hashCode.AddBytes(Value);
        return hashCode.ToHashCode();
    }

    public static readonly BorderStyle None = new("none"u8.ToArray());

    public static readonly BorderStyle Thin = new("thin"u8.ToArray());

    public static readonly BorderStyle Medium = new("medium"u8.ToArray());

    public static readonly BorderStyle Dashed = new("dashed"u8.ToArray());

    public static readonly BorderStyle Dotted = new("dotted"u8.ToArray());

    public static readonly BorderStyle Thick = new("thick"u8.ToArray());

    public static readonly BorderStyle Double = new("double"u8.ToArray());

    public static readonly BorderStyle Hair = new("hair"u8.ToArray());

    public static readonly BorderStyle MediumDashed = new("mediumDashed"u8.ToArray());

    public static readonly BorderStyle DashDot = new("dashDot"u8.ToArray());

    public static readonly BorderStyle MediumDashDot = new("mediumDashDot"u8.ToArray());

    public static readonly BorderStyle DashDotDot = new("dashDotDot"u8.ToArray());

    public static readonly BorderStyle MediumDashDotDot = new("mediumDashDotDot"u8.ToArray());

    public static readonly BorderStyle SlantDashDot = new("slantDashDot"u8.ToArray());
}
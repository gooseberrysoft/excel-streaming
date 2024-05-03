using System.Text;

namespace Gooseberry.ExcelStreaming.Styles;

public readonly struct FillPattern : IEquatable<FillPattern>
{
    private readonly byte[] _value;

    internal FillPattern(byte[] value)
    {
        _value = value;
    }

    internal byte[] Value => _value ?? None.Value;

    public override string ToString()
        => Encoding.UTF8.GetString(Value);

    public bool Equals(FillPattern other)
        => Value.AsSpan().SequenceEqual(other.Value);

    public override bool Equals(object? obj)
        => obj is FillPattern other && Equals(other);

    public override int GetHashCode()
    {
        var hashCode = new HashCode();
        hashCode.AddBytes(Value);
        return hashCode.ToHashCode();
    }

    public static readonly FillPattern None = new("none"u8.ToArray());

    public static readonly FillPattern Solid = new("solid"u8.ToArray());

    public static readonly FillPattern MediumGray = new("mediumGray"u8.ToArray());

    public static readonly FillPattern DarkGray = new("darkGray"u8.ToArray());

    public static readonly FillPattern LightGray = new("lightGray"u8.ToArray());

    public static readonly FillPattern DarkHorizontal = new("darkHorizontal"u8.ToArray());

    public static readonly FillPattern DarkVertical = new("darkVertical"u8.ToArray());

    public static readonly FillPattern DarkDown = new("darkDown"u8.ToArray());

    public static readonly FillPattern DarkUp = new("darkUp"u8.ToArray());

    public static readonly FillPattern DarkGrid = new("darkGrid"u8.ToArray());

    public static readonly FillPattern DarkTrellis = new("darkTrellis"u8.ToArray());

    public static readonly FillPattern LightHorizontal = new("lightHorizontal"u8.ToArray());

    public static readonly FillPattern LightVertical = new("lightVertical"u8.ToArray());

    public static readonly FillPattern LightDown = new("lightDown"u8.ToArray());

    public static readonly FillPattern LightUp = new("lightUp"u8.ToArray());

    public static readonly FillPattern LightGrid = new("lightGrid"u8.ToArray());

    public static readonly FillPattern LightTrellis = new("lightTrellis"u8.ToArray());

    public static readonly FillPattern Gray125 = new("gray125"u8.ToArray());

    public static readonly FillPattern Gray0625 = new("gray0625"u8.ToArray());
}
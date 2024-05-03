using System.Text;

namespace Gooseberry.ExcelStreaming.Styles;

public readonly struct HorizontalAlignment : IEquatable<HorizontalAlignment>
{
    private readonly byte[] _value;

    internal HorizontalAlignment(byte[] value)
    {
        _value = value;
    }

    internal byte[] Value => _value ?? General.Value;

    public override string ToString()
        => Encoding.UTF8.GetString(Value);

    public bool Equals(HorizontalAlignment other)
        => Value.AsSpan().SequenceEqual(other.Value);

    public override bool Equals(object? obj)
        => obj is HorizontalAlignment other && Equals(other);

    public override int GetHashCode()
    {
        var hashCode = new HashCode();
        hashCode.AddBytes(Value);
        return hashCode.ToHashCode();
    }

    public static readonly HorizontalAlignment General = new("general"u8.ToArray());

    public static readonly HorizontalAlignment Left = new("left"u8.ToArray());

    public static readonly HorizontalAlignment Center = new("center"u8.ToArray());

    public static readonly HorizontalAlignment Right = new("right"u8.ToArray());

    public static readonly HorizontalAlignment Fill = new("fill"u8.ToArray());

    public static readonly HorizontalAlignment Justify = new("justify"u8.ToArray());

    public static readonly HorizontalAlignment CenterContinuous = new("centerContinuous"u8.ToArray());

    public static readonly HorizontalAlignment Distributed = new("distributed"u8.ToArray());
}
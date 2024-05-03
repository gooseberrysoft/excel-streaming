using System.Text;

namespace Gooseberry.ExcelStreaming.Styles;

public readonly struct VerticalAlignment : IEquatable<VerticalAlignment>
{
    private readonly byte[] _value;

    internal VerticalAlignment(byte[] value)
        => _value = value;

    internal byte[] Value => _value ?? Top.Value;

    public override string ToString()
        => Encoding.UTF8.GetString(Value);

    public bool Equals(VerticalAlignment other)
        => Value.AsSpan().SequenceEqual(other.Value);

    public override bool Equals(object? obj)
        => obj is VerticalAlignment other && Equals(other);

    public override int GetHashCode()
    {
        var hashCode = new HashCode();
        hashCode.AddBytes(Value);
        return hashCode.ToHashCode();
    }

    public static readonly VerticalAlignment Top = new("top"u8.ToArray());

    public static readonly VerticalAlignment Center = new("center"u8.ToArray());

    public static readonly VerticalAlignment Bottom = new("bottom"u8.ToArray());

    public static readonly VerticalAlignment Justify = new("justify"u8.ToArray());

    public static readonly VerticalAlignment Distributed = new("distributed"u8.ToArray());
}
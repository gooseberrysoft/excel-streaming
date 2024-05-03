using System.Text;

namespace Gooseberry.ExcelStreaming.Styles;

public readonly struct Underline : IEquatable<Underline>
{
    private readonly byte[] _value;

    internal Underline(byte[] value)
        => _value = value;

    internal byte[] Value => _value ?? None.Value;

    public override string ToString()
        => Encoding.UTF8.GetString(Value);

    public bool Equals(Underline other)
        => Value.AsSpan().SequenceEqual(other.Value);

    public override bool Equals(object? obj)
        => obj is Underline other && Equals(other);

    public override int GetHashCode()
    {
        var hashCode = new HashCode();
        hashCode.AddBytes(Value);
        return hashCode.ToHashCode();
    }

    public static Underline Single => new ("single"u8.ToArray());

    public static Underline Double => new ("double"u8.ToArray());

    public static Underline SingleAccounting => new ("singleAccounting"u8.ToArray());

    public static Underline DoubleAccounting => new ("doubleAccounting"u8.ToArray());

    public static Underline None => new ("none"u8.ToArray());
}
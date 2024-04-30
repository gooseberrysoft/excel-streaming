namespace Gooseberry.ExcelStreaming.Styles;

public readonly struct StyleReference : IEquatable<StyleReference>
{
    internal StyleReference(int value)
        => Value = value;

    internal int Value { get; }

    public bool Equals(StyleReference other)
        => Value == other.Value;

    public override bool Equals(object? obj)
        => obj is StyleReference other && Equals(other);

    public override int GetHashCode()
        => Value;
}
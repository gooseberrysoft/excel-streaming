// ReSharper disable once CheckNamespace
namespace Gooseberry.ExcelStreaming;

public readonly struct SharedStringReference : IStringReferences<SharedStringReference, string>
{
    internal SharedStringReference(int value)
    {
        Value = value;
    }

    internal int Value { get; }

    void IStringReferences<SharedStringReference, string>.SetStrings(SharedStringList sharedStrings, string value)
    {
        sharedStrings[this] = value;
    }

    public static SharedStringReference Create(SharedStringList sharedStrings, string? defaultValue)
    {
        var reference = sharedStrings.GetNextReference();

        if (defaultValue != null)
            sharedStrings[reference] = defaultValue;

        return reference;
    }
}
// ReSharper disable once CheckNamespace

namespace Gooseberry.ExcelStreaming;

public readonly struct SharedStringReference
{
    internal SharedStringReference(int value)
    {
        Value = value;
    }

    internal int Value { get; }
}
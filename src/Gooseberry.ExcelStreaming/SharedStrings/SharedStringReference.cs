// ReSharper disable once CheckNamespace

namespace Gooseberry.ExcelStreaming;

public readonly struct SharedStringReference(int value)
{
    internal int Value { get; } = value;
}
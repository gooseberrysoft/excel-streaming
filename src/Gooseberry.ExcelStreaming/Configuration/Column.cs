// ReSharper disable once CheckNamespace

namespace Gooseberry.ExcelStreaming;

public readonly record struct Column(decimal Width, bool IsHidden = false)
{
    public Column(decimal width) : this(width, false)
    {
    }
}
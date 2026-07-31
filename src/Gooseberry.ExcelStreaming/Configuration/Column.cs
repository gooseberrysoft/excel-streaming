// ReSharper disable once CheckNamespace

namespace Gooseberry.ExcelStreaming;

public readonly record struct Column(decimal Width, bool IsHidden = false)
{
    public Column(decimal Width) : this(Width, false)
    {
    }

    public void Deconstruct(out decimal Width) => Width = this.Width;
}
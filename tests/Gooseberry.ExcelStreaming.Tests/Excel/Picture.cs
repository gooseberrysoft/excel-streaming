using Gooseberry.ExcelStreaming.Pictures.Abstractions;
using Gooseberry.ExcelStreaming.Tests.Extensions;

namespace Gooseberry.ExcelStreaming.Tests.Excel;

public sealed record Picture(
    byte[] Data,
    IPicturePlacement Placement,
    PictureFormat Format)
{
    public override int GetHashCode()
        => HashCode.Combine(Placement, Format, Data.GetCollectionHashCode());

    public bool Equals(Picture? other)
    {
        if (other is null)
            return false;

        return Placement == other.Placement
            && Format == other.Format
            && Data.SequenceEqual(other.Data);
    }
}

using System.Drawing;
using Gooseberry.ExcelStreaming.Tests.Extensions;

namespace Gooseberry.ExcelStreaming.Tests.Excel;

public readonly record struct PicturePlacement(AnchorCell From, AnchorCell? To, Size? Size)
{
    public PicturePlacement(AnchorCell from, AnchorCell to)
        : this(from, to, null)
    {
    }

    public PicturePlacement(AnchorCell from, Size size)
        : this(from, null, size)
    {
    }
};

public sealed record Picture(
    byte[] Data,
    PicturePlacement Placement,
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
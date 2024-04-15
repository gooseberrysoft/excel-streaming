using System.Drawing;
using Gooseberry.ExcelStreaming.Pictures.Abstractions;

namespace Gooseberry.ExcelStreaming.Pictures.Placements;

public sealed record OneCellAnchorPicturePlacement(AnchorCell From, Size Size) : IPicturePlacement
{
    public void Visit<T>(T visitor) where T : IPicturePlacementVisitor
        => visitor.Visit(this);
}

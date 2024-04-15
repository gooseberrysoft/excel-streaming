using Gooseberry.ExcelStreaming.Pictures.Abstractions;

namespace Gooseberry.ExcelStreaming.Pictures.Placements;

public sealed record TwoCellAnchorPicturePlacement(AnchorCell From, AnchorCell To) : IPicturePlacement
{
    public void Visit<T>(T visitor) where T : IPicturePlacementVisitor
        => visitor.Visit(this);
}

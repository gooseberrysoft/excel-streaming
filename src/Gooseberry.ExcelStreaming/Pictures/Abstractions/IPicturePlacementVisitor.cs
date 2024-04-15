using Gooseberry.ExcelStreaming.Pictures.Placements;

namespace Gooseberry.ExcelStreaming.Pictures.Abstractions;

public interface IPicturePlacementVisitor
{
    void Visit(OneCellAnchorPicturePlacement placement);
    void Visit(TwoCellAnchorPicturePlacement placement);
}

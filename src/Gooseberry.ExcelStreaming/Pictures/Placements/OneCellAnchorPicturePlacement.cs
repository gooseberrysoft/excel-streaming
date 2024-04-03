using System.Drawing;
using Gooseberry.ExcelStreaming.Pictures.Abstractions;

namespace Gooseberry.ExcelStreaming.Pictures.Placements;

public sealed record OneCellAnchorPicturePlacement(AnchorCell From, Size Size) : IPicturePlacement;

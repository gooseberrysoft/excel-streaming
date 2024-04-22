using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.Pictures;

internal sealed record Picture(
    int Id,
    string RelationshipId,
    string Name,
    PictureData Data,
    PictureFormat Format,
    IPicturePlacementWriter PlacementWriter);
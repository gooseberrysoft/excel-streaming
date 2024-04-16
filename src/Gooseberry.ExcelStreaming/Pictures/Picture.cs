using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.Pictures;

internal readonly record struct Picture(
    int Id,
    string RelationshipId,
    string Name,
    PictureData Data,
    PictureFormat Format,
    IPicturePlacementWriter PlacementWriter);
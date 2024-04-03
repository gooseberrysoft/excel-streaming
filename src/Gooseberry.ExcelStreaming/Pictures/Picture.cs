using Gooseberry.ExcelStreaming.Pictures.Abstractions;

namespace Gooseberry.ExcelStreaming.Pictures;

internal readonly record struct Picture(
    int Id,
    string RelationshipId,
    string Name,
    PictureData Data,
    IPicturePlacement Placement);

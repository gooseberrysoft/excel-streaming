using Gooseberry.ExcelStreaming.Pictures.Abstractions;

namespace Gooseberry.ExcelStreaming.Pictures;

internal readonly record struct Picture(
    int Id,
    string RelationshipId,
    string Name,
    PictureData Data,
    PictureInfo Info,
    IPicturePlacement Placement)
{
    public Picture(
        int Id,
        string RelationshipId,
        string Name,
        PictureData Data,
        IPicturePlacement Placement)
        : this(Id, RelationshipId, Name, Data, Data.GetInfo(), Placement)
    {
    }
}

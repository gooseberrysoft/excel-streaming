using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.Pictures;

internal sealed class SheetDrawings
{
    private readonly Dictionary<int, Drawing> _drawings = new();

    private int _id = 1;

    public IEnumerable<Picture> Pictures => _drawings.Values.SelectMany(v => v.Pictures);

    public Drawing Get(int sheetId)
    {
        if (!_drawings.TryGetValue(sheetId, out var pictures))
        {
            _drawings[sheetId] = pictures = new Drawing(sheetId);
        }

        return pictures;
    }

    public void AddPicture(int sheetId, in PictureData data, PictureFormat format, IPicturePlacementWriter placementWriter)
    {
        var id = _id++;
        var relationshipId = $"rId{id}";
        var name = $"Image{id}";
        var picture = new Picture(id, RelationshipId: relationshipId, Name: name, data, format, placementWriter);

        Get(sheetId).Add(picture);
    }
}
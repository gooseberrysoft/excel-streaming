using Gooseberry.ExcelStreaming.Pictures.Abstractions;

namespace Gooseberry.ExcelStreaming.Pictures;

internal sealed class ExcelPictures
{
    private readonly Dictionary<int, Drawing> _sheetPictures = new();

    private int _id = 1;

    public IEnumerable<Picture> Pictures => _sheetPictures.Values.SelectMany(v => v.Pictures);

    public Drawing Get(int sheetId)
    {
        if (_sheetPictures.TryGetValue(sheetId, out var pictures))
            return pictures;

        throw new InvalidOperationException($"Unable to find pictures for sheet {sheetId}.");
    }

    public void AddPicture(int sheetId, PictureData data, IPicturePlacement placement)
    {
        var id = _id++;
        var relationshipId = $"rId{id}";
        var name = $"Picture {id}";
        var picture = new Picture(id, RelationshipId: relationshipId, Name: name, data, placement);

        if (!_sheetPictures.TryGetValue(sheetId, out var pictures))
            _sheetPictures[sheetId] = pictures = new Drawing(sheetId);

        pictures.Add(picture);
    }
}

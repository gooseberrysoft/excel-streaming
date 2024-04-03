namespace Gooseberry.ExcelStreaming.Pictures;

internal sealed class Drawing
{
    private readonly List<Picture> _pictures;

    public IReadOnlyCollection<Picture> Pictures => _pictures;

    public int SheetId { get; }
    
    public string RelationshipId { get; }

    public bool IsEmpty => Pictures.Count == 0;

    public Drawing(int sheetId)
    {
        SheetId = sheetId;
        RelationshipId = $"dId{sheetId}";
        _pictures = new List<Picture>();
    }

    public void Add(Picture picture)
        => _pictures.Add(picture);
}

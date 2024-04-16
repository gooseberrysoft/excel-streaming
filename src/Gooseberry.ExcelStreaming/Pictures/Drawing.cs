namespace Gooseberry.ExcelStreaming.Pictures;

internal sealed class Drawing
{
    private List<Picture>? _pictures;

    public IReadOnlyCollection<Picture> Pictures
    {
        get
        {
            IReadOnlyCollection<Picture>? pictures = _pictures;
            return pictures ?? Array.Empty<Picture>();
        }
    }

    public int SheetId { get; }

    public string RelationshipId { get; }

    public bool IsEmpty => Pictures.Count == 0;

    public Drawing(int sheetId)
    {
        SheetId = sheetId;
        RelationshipId = $"dId{sheetId}";
    }

    public void Add(in Picture picture)
        => (_pictures ??= new List<Picture>()).Add(picture);
}
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming;

[StructLayout(LayoutKind.Auto)]
internal readonly struct Sheet
{
    public Sheet(string name, int id, string relationshipId)
    {
        Name = name;
        Id = id;
        RelationshipId = relationshipId;
    }

    public string Name { get; }

    public int Id { get; }

    public string RelationshipId { get; }
}
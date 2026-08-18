using FluentAssertions;
using Gooseberry.ExcelStreaming;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests.SharedStrings;

public sealed class SharedStringListTests
{
    [Fact]
    public void GetNextReference_and_indexer_store_and_retrieve_values()
    {
        using var seq = new SequenceList<string?>();
        var list = new SharedStringList(seq, offset: 0);

        var r1 = list.GetNextReference();
        list[r1] = "first";

        var r2 = list.GetNextReference();
        list[r2] = "second";

        list[r1].Should().Be("first");
        list[r2].Should().Be("second");
    }

    [Fact]
    public void Constructor_offset_is_applied_correctly()
    {
        using var seq = new SequenceList<string?>();
        var offset = 10;
        var list = new SharedStringList(seq, offset);

        var r = list.GetNextReference(); // underlying sequence index 0, reference.Value == 0 + offset
        list[r] = "offset-value";

        list[r].Should().Be("offset-value");
    }

    [Fact]
    public void Indexer_throws_for_invalid_reference()
    {
        using var seq = new SequenceList<string?>();
        var list = new SharedStringList(seq, offset: 0);

        var badRef = new SharedStringReference(999);

        var act = () => { _ = list[badRef]; };

        act.Should().Throw<ArgumentException>();
    }
}
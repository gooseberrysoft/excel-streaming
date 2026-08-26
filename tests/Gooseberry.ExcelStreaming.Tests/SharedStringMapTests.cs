using FluentAssertions;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class SharedStringMapTests
{
    [Fact]
    public void Add_ReturnsSameReference_ForSameKey()
    {
        // Arrange
        var sharedStrings = new SharedStringList(new SequenceList<string?>(), 0);
        var map = new SharedStringMap<string>(sharedStrings, k => "value");

        // Act
        var first = map.GetOrAdd("key");
        var second = map.GetOrAdd("key");

        // Assert
        first.Should().Be(second);

        sharedStrings[first].Should().Be("value");
        sharedStrings[second].Should().Be("value");
    }

    [Fact]
    public void Add_ReturnsDifferentReferences_ForDifferentKeys()
    {
        // Arrange
        var sharedStrings = new SharedStringList(new SequenceList<string?>(), 0);
        var map = new SharedStringMap<string>(sharedStrings,
            new Dictionary<string, string> { { "first", "value1" }, { "second", "value2" } });

        // Act
        var first = map.GetOrAdd("first");
        var second = map.GetOrAdd("second");

        // Assert
        first.Should().NotBe(second);

        sharedStrings[first].Should().Be("value1");
        sharedStrings[second].Should().Be("value2");
    }
}
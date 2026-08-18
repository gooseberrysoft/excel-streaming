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
        var map = new SharedStringMap<string>(sharedStrings);

        // Act
        var first = map.Add("key", _ => "value");
        var second = map.Add("key", _ => "another value");

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
        var map = new SharedStringMap<string>(sharedStrings);

        // Act
        var first = map.Add("first", _ => "value1");
        var second = map.Add("second", _ => "value2");

        // Assert
        first.Should().NotBe(second);

        sharedStrings[first].Should().Be("value1");
        sharedStrings[second].Should().Be("value2");
    }
}
using FluentAssertions;
using NSubstitute;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class SharedStringResolverTests
{
    [Fact]
    public void Add_ReturnsSameReference_ForDuplicateKey()
    {
        var sharedStringList = new SharedStringList(new SequenceList<string?>(), 0);

        var provider = Substitute.For<IStringProvider<string>>();

        var resolver = new SharedStringResolver<string>(providerBatchSize: 10, sharedStringList, provider);

        var r1 = resolver.Add("key");
        var r2 = resolver.Add("key");

        r1.Value.Should().Be(r2.Value);
    }

    [Fact]
    public async Task FlushBatch_LoadsStrings_FromProvider()
    {
        var sharedStringList = new SharedStringList(new SequenceList<string?>(), 0);

        var provider = Substitute.For<IStringProvider<string>>();

        // When provider receives the batch ["a","b"] return mapped values
        provider.GetStrings(Arg.Any<IReadOnlyCollection<string>>())
            .Returns(call =>
            {
                var keys = call.Arg<IReadOnlyCollection<string>>();
                var result = keys!.Select(k => KeyValuePair.Create(k, $"val-{k}")).ToArray();
                return Task.FromResult<IEnumerable<KeyValuePair<string, string>>>(result);
            });

        var resolver = new SharedStringResolver<string>(providerBatchSize: 2, sharedStringList, provider);

        var refA = resolver.Add("a");
        var refB = resolver.Add("b");

        await resolver.Complete();

        sharedStringList[refA].Should().Be("val-a");
        sharedStringList[refB].Should().Be("val-b");
    }

    [Fact]
    public async Task FlushSeveralBatches_LoadsStrings_FromProvider()
    {
        var sharedStringList = new SharedStringList(new SequenceList<string?>(), 0);

        var provider = Substitute.For<IStringProvider<int>>();

        provider.GetStrings(Arg.Any<IReadOnlyCollection<int>>())
            .Returns(call =>
            {
                var keys = call.Arg<IReadOnlyCollection<int>>();
                var result = keys!.Select(k => KeyValuePair.Create(k, $"val-{k}")).ToArray();
                return Task.FromResult<IEnumerable<KeyValuePair<int, string>>>(result);
            });

        var resolver = new SharedStringResolver<int>(providerBatchSize: 2, sharedStringList, provider);

        var refA = resolver.Add(1);
        var refB = resolver.Add(2);
        var refC = resolver.Add(3);

        await resolver.Complete();

        sharedStringList[refA].Should().Be("val-1");
        sharedStringList[refB].Should().Be("val-2");
        sharedStringList[refC].Should().Be("val-3");
    }

    [Theory]
    [InlineData("DEFAULT")]
    [InlineData(null)]
    public async Task Complete_AppliesDefaultValue_WhenProviderMissingEntries(string? defaultValue)
    {
        var sharedStringList = new SharedStringList(new SequenceList<string?>(), 0);

        var provider = Substitute.For<IStringProvider<string>>();

        // Provider returns only entry for "present" key, omits "missing"
        provider.GetStrings(Arg.Any<IReadOnlyCollection<string>>())
            .Returns(call =>
            {
                var keys = call.Arg<IReadOnlyCollection<string>>();
                var result = keys!
                    .Where(k => k == "present")
                    .Select(k => KeyValuePair.Create(k, "present-value"))
                    .ToArray();
                return Task.FromResult<IEnumerable<KeyValuePair<string, string>>>(result);
            });

        var resolver = new SharedStringResolver<string>(providerBatchSize: 2, sharedStringList, provider, defaultValue);

        var refPresent = resolver.Add("present");
        var refMissing = resolver.Add("missing");

        await resolver.Complete();

        sharedStringList[refPresent].Should().Be("present-value");
        sharedStringList[refMissing].Should().Be(defaultValue);
    }
}
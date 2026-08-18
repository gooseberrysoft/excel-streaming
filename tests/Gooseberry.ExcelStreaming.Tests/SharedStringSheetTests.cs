using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using FluentAssertions;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class SharedStringSheetTests
{
    [Fact]
    public void GetOrAdd_ReturnsSameReference_And_ItemsAccessible()
    {
        var sheet = new SharedStringSheet(sharedStringTable: null);

        var ref1 = sheet.GetOrAdd("hello");
        var ref2 = sheet.GetOrAdd("hello");
        var ref3 = sheet.GetOrAdd("world");

        // same string -> same reference
        ref1.Value.Should().Be(ref2.Value);
        // different string -> different reference
        ref3.Value.Should().NotBe(ref1.Value);

        sheet.Items[ref1].Should().Be("hello");
        sheet.Items[ref3].Should().Be("world");
    }

    [Fact]
    public async Task WriteTo_NoSharedTableAndNoStrings_WritesEmptyTableUsingWriteEntry()
    {
        var sheet = new SharedStringSheet(sharedStringTable: null);

        var buffer = new BufferSequence(Buffer.MinSize);
        var encoder = Encoding.UTF8.GetEncoder();

        var archive = new FakeArchiveWriter();

        await sheet.WriteTo(buffer, encoder, archive, "xl/sharedStrings.xml");

        archive.WrittenEntries.Should().ContainSingle();
        var (path, mem) = archive.WrittenEntries[0];
        path.Should().Be("xl/sharedStrings.xml");

        var expected = Gooseberry.ExcelStreaming.Writers.SharedStringWriter.EmptyTable;
        mem.ToArray().Should().Equal(expected);
    }

    [Fact]
    public async Task WriteTo_WithInlineStrings_WritesStringsToCreatedEntry()
    {
        var sheet = new SharedStringSheet(sharedStringTable: null);

        // add inline strings
        sheet.GetOrAdd("one & two");
        sheet.GetOrAdd("three < four");

        var buffer = new BufferSequence(Buffer.MinSize);
        var encoder = Encoding.UTF8.GetEncoder();

        var archive = new FakeArchiveWriter();

        await sheet.WriteTo(buffer, encoder, archive, "xl/sharedStrings.xml");

        // verify that an entry was created and written to
        archive.CreatedEntryPaths.Should().ContainSingle("xl/sharedStrings.xml");

        var content = Encoding.UTF8.GetString(archive.EntryStream.ToArray());

        // basic checks: prefix, suffix and presence of escaped values
        content.Should().StartWith("<sst ");
        content.Should().EndWith("</sst>");
        content.Should().Contain("<t>one &amp; two</t>").And.Contain("<t>three &lt; four</t>");
    }

    [Fact]
    public void WriteTo_BufferNotEmpty_ThrowsArgumentException()
    {
        var sheet = new SharedStringSheet(sharedStringTable: null);
        sheet.GetOrAdd("one & two");
        
        var buffer = new BufferSequence(Buffer.MinSize);
        var encoder = Encoding.UTF8.GetEncoder();

        // make buffer non-empty
        var span = buffer.GetSpan();
        span[0] = (byte)'x';
        buffer.Advance(1);

        var archive = new FakeArchiveWriter();

        Func<ValueTask> act = () => sheet.WriteTo(buffer, encoder, archive, "xl/sharedStrings.xml");

        act.Should().Throw<ArgumentException>().WithMessage("Buffer must be empty");
    }

    private sealed class FakeArchiveWriter : IArchiveWriter
    {
        public readonly List<(string Path, ReadOnlyMemory<byte> Buffer)> WrittenEntries = new();
        public readonly MemoryStream EntryStream = new();
        public readonly List<string> CreatedEntryPaths = new();

        public ValueTask WriteEntry(string entryPath, ReadOnlyMemory<byte> buffer)
        {
            WrittenEntries.Add((entryPath, buffer));
            return ValueTask.CompletedTask;
        }

        public async ValueTask WriteEntry(string entryPath, Stream stream)
        {
            CreatedEntryPaths.Add(entryPath);
            using var ms = new MemoryStream();
            await stream.CopyToAsync(ms);
            WrittenEntries.Add((entryPath, ms.ToArray()));
        }

        public IEntryWriter CreateEntry(string entryPath)
        {
            CreatedEntryPaths.Add(entryPath);
            EntryStream.SetLength(0);
            return new EntryWriter(EntryStream);
        }

        public ValueTask DisposeAsync() => ValueTask.CompletedTask;

        private sealed class EntryWriter : IEntryWriter
        {
            private readonly MemoryStream _stream;

            public EntryWriter(MemoryStream stream) => _stream = stream;

            public ValueTask Write(MemoryOwner buffer)
            {
                _stream.Write(buffer.Memory.Span);
                return ValueTask.CompletedTask;
            }

            public bool TryWrite(in MemoryOwner buffer)
            {
                _stream.Write(buffer.Memory.Span);
                return true;
            }
        }
    }
}
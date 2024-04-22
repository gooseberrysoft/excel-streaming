using System.IO;
using System.Threading.Tasks;
using FluentAssertions;
using Gooseberry.ExcelStreaming.SharedStrings;
using Gooseberry.ExcelStreaming.Tests.Excel;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class ExcelWriterSharedStringTableTests
{
    [Fact]
    public async Task ExcelWriter_WritesCorrectSharedStrings()
    {
        var outputStream = new MemoryStream();

        var builder = new SharedStringTableBuilder();
        var stringReference = builder.GetOrAdd("string");
        var otherStringReference = builder.GetOrAdd("other string");
                
        await using (var writer = new ExcelWriter(outputStream, sharedStringTable: builder.Build()))
        {
            await writer.StartSheet("test sheet");
            await writer.StartRow();
            
            writer.AddCell(stringReference);
            writer.AddCell(otherStringReference);
            
            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);
        var sharedStrings = ExcelReader.ReadSharedStrings(outputStream);

        sharedStrings.Should().BeEquivalentTo(new[] {"string", "other string"});
        
        outputStream.Seek(0, SeekOrigin.Begin);
        var sheets = ExcelReader.ReadSheets(outputStream);
        
        var expectedSheet = new Excel.Sheet(
            "test sheet",
            new []
            {
                new Row(new []
                {
                    new Cell("0", CellValueType.SharedString),
                    new Cell("1", CellValueType.SharedString),
                })
            });

        sheets.ShouldBeEquivalentTo(expectedSheet);        
    }
    
    [Fact]
    public async Task ExcelWriter_AddCellWithSharedStringAndSharedStringTable_WritesCorrect()
    {
        var outputStream = new MemoryStream();

        var builder = new SharedStringTableBuilder();
        var stringReference = builder.GetOrAdd("string");
        var otherStringReference = builder.GetOrAdd("other string");
                
        await using (var writer = new ExcelWriter(outputStream, sharedStringTable: builder.Build()))
        {
            await writer.StartSheet("test sheet");
            await writer.StartRow();
            
            writer.AddCell(stringReference);
            writer.AddCell(otherStringReference);
            writer.AddCellWithSharedString("third string");
            writer.AddCellWithSharedString("one more string");
            
            await writer.Complete();
        }

        outputStream.Seek(0, SeekOrigin.Begin);
        var sharedStrings = ExcelReader.ReadSharedStrings(outputStream);

        sharedStrings.Should().BeEquivalentTo(new[] {"string", "other string", "third string", "one more string"});
        
        outputStream.Seek(0, SeekOrigin.Begin);
        var sheets = ExcelReader.ReadSheets(outputStream);
        
        var expectedSheet = new Excel.Sheet(
            "test sheet",
            new []
            {
                new Row(new []
                {
                    new Cell("0", CellValueType.SharedString),
                    new Cell("1", CellValueType.SharedString),
                    new Cell("2", CellValueType.SharedString),
                    new Cell("3", CellValueType.SharedString),
                })
            });

        sheets.ShouldBeEquivalentTo(expectedSheet);        
    }    
}

using FluentAssertions;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class ExcelWriterErrorTests : IAsyncLifetime
{
    private readonly ExcelWriter _excelWriter;

    public ExcelWriterErrorTests()
        => _excelWriter = new ExcelWriter(new MemoryStream());

    [Fact]
    public void CreateWithNullStream_ThrowsException()
    {
        Action action = () => new ExcelWriter((Stream)null!);

        action.Should().ThrowExactly<ArgumentNullException>();
    }

    [Fact]
    public async Task StartRowWithoutStartSheet_ThrowsException()
    {
        Func<Task> action = () => _excelWriter.StartRow().AsTask();

        await action.Should()
            .ThrowExactlyAsync<InvalidOperationException>()
            .WithMessage("Sheet is not started.");
    }

    
    [Fact]
    public async Task AddCellWithoutStartRow_ThrowsException()
    {
        await _excelWriter.StartSheet("test");

        CheckAddCell(() => _excelWriter.AddCell("test"));
        CheckAddCell(() => _excelWriter.AddCell(1));
        CheckAddCell(() => _excelWriter.AddCell(1L));
        CheckAddCell(() => _excelWriter.AddCell(1.0M));
        CheckAddCell(() => _excelWriter.AddCell(DateTime.Now));
        CheckAddCell(() => _excelWriter.AddEmptyCell());
        CheckAddCell(() => _excelWriter.AddCell((int?)null));
        CheckAddCell(() => _excelWriter.AddCell((long?)null));
        CheckAddCell(() => _excelWriter.AddCell((decimal?)null));
        CheckAddCell(() => _excelWriter.AddCell((DateTime?)null));

        void CheckAddCell(Action action)
        {
            action.Should()
                .ThrowExactly<InvalidOperationException>()
                .WithMessage("Row is not started yet.");
        }
    }

    [Fact]
    public async Task AfterComplete_ThrowsException()
    {
        await _excelWriter.Complete();

        await CheckAction(() => _excelWriter.StartSheet("test"));
        await CheckAction(() => _excelWriter.StartRow());
        //CheckAddCell(() => _excelWriter.AddCell("test"));
        //CheckAddCell(() => _excelWriter.AddCell(1));
        //CheckAddCell(() => _excelWriter.AddCell(1L));
        //CheckAddCell(() => _excelWriter.AddCell(1.0M));
        //CheckAddCell(() => _excelWriter.AddCell(DateTime.Now));
        //CheckAddCell(() => _excelWriter.AddEmptyCell());
        //CheckAddCell(() => _excelWriter.AddCell((int?)null));
        //CheckAddCell(() => _excelWriter.AddCell((long?)null));
        //CheckAddCell(() => _excelWriter.AddCell((decimal?)null));
        //CheckAddCell(() => _excelWriter.AddCell((DateTime?)null));
        await CheckAction(() => _excelWriter.Complete());

        async Task CheckAction(Func<ValueTask> action)
        {
            Func<Task> actionToCheck = () => action().AsTask();
            await actionToCheck.Should()
                .ThrowExactlyAsync<InvalidOperationException>()
                .WithMessage("Excel writer is already completed.");
        }

        //void CheckAddCell(Action action)
        //{
        //    action.Should()
        //        .ThrowExactly<InvalidOperationException>()
        //        .WithMessage("Excel writer is already completed.");
        //}
    }

    public Task InitializeAsync()
        => Task.CompletedTask;

    public async Task DisposeAsync()
        => await _excelWriter.DisposeAsync();
}
using System;
using System.IO;
using System.Threading.Tasks;
using FluentAssertions;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests
{
    public sealed class ExcelWriterErrorTests : IAsyncLifetime
    {
        private readonly ExcelWriter _excelWriter;

        public ExcelWriterErrorTests()
            => _excelWriter = new ExcelWriter(new MemoryStream());

        [Fact]
        public void CreateWithNullStream_ThrowsException()
        {
            Action action = () => new ExcelWriter(null!);

            action.Should().ThrowExactly<ArgumentNullException>();
        }

        [Fact]
        public async Task StartRowWithoutStartSheet_ThrowsException()
        {
            Func<Task> action = () => _excelWriter.StartRow().AsTask();

            await action.Should()
                .ThrowExactlyAsync<InvalidOperationException>()
                .WithMessage("Cannot start row before start sheet.");
        }

        [Theory]
        [InlineData(-10.0)]
        [InlineData(-0.1)]
        [InlineData(0)]
        public async Task StartRowIncorrectHeight_ThrowsException(double height)
        {
            await _excelWriter.StartSheet("test");
            Func<Task> action = () => _excelWriter.StartRow((decimal)height).AsTask();

            await action.Should()
                .ThrowExactlyAsync<ArgumentOutOfRangeException>()
                .WithMessage("Height of row cannot be less than 0. (Parameter 'height')");
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
                    .WithMessage("Cannot add cell before start row.");
            }
        }

        [Fact]
        public async Task AnyActionAfterComplete_ThrowsException()
        {
            await _excelWriter.Complete();

            await CheckAction(() => _excelWriter.StartSheet("test"));
            await CheckAction(() => _excelWriter.StartRow());
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
            await CheckAction(() => _excelWriter.Complete());

            async Task CheckAction(Func<ValueTask> action)
            {
                Func<Task> actionToCheck = () => action().AsTask();
                await actionToCheck.Should()
                    .ThrowExactlyAsync<InvalidOperationException>()
                    .WithMessage("Cannot use excel writer. It is completed already.");
            }

            void CheckAddCell(Action action)
            {
                action.Should()
                    .ThrowExactly<InvalidOperationException>()
                    .WithMessage("Cannot use excel writer. It is completed already.");
            }
        }

        public Task InitializeAsync()
            => Task.CompletedTask;

        public async Task DisposeAsync()
            => await _excelWriter.DisposeAsync();
    }
}

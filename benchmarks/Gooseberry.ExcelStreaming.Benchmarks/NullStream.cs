namespace Gooseberry.ExcelStreaming.Benchmarks
{
    public sealed class NullStream : Stream
    {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => 0;
        public override long Position
        {
            get => 0;
            set { }
        }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count) => 0;

        public override long Seek(long offset, SeekOrigin origin) => 0;

        public override void SetLength(long value)
        {
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
        }

        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default) 
            => ValueTask.CompletedTask;

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) 
            => Task.CompletedTask;

        public override Task FlushAsync(CancellationToken cancellationToken) 
            => Task.CompletedTask;
    }
}
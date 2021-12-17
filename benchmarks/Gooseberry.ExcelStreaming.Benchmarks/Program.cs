using BenchmarkDotNet.Running;

namespace Gooseberry.ExcelStreaming.Benchmarks
{
    class Program
    {
        static void Main(string[] args) 
            => BenchmarkRunner.Run<ExcelWriterBenchmarks>();
    }
}
using BenchmarkDotNet.Running;

namespace Gooseberry.ExcelStreaming.Benchmarks
{
    class Program
    {
        static void Main(string[] args) 
            => BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
    }
}
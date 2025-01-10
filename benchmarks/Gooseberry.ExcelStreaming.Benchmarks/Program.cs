using BenchmarkDotNet.Running;

namespace Gooseberry.ExcelStreaming.Benchmarks;

class Program
{
    static async Task Main(string[] args)
    {
        if (args.Length > 0 && args[0] == "--profilerMode")
        {
            var benchmark = new RealWorldReportBenchmarks();
            benchmark.RowsCount = args.Length > 1 && int.TryParse(args[1], out var result) ? result : 100_000;

            while (true)
            {
                await benchmark.RealWorldReport();
            }
        }
        else
            BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
    }
}
using BenchmarkDotNet.Running;

namespace LargeXlsx.Benchmarks
{
    internal class Program
    {
        static void Main(string[] args)
        {
            BenchmarkRunner.Run<WriterLowLevelBenchmarks>();
        }
    }
}

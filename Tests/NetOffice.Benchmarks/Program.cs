using System;
using BenchmarkDotNet.Running;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Exporters;
using BenchmarkDotNet.Exporters.Csv;

namespace NetOffice.Benchmarks
{
    /// <summary>
    /// Entry point for NetOffice.Benchmarks
    /// Benchmarks the performance of RemoveObjectFromList with different collection implementations
    /// to address Issue #221: https://github.com/NetOfficeFw/NetOffice/issues/221
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=======================================================");
            Console.WriteLine("NetOffice RemoveObjectFromList Performance Benchmarks");
            Console.WriteLine("Issue #221: Core.RemoveObjectFromList Performance");
            Console.WriteLine("=======================================================");
            Console.WriteLine();
            Console.WriteLine("This benchmark suite compares four collection implementations:");
            Console.WriteLine("  1. List<T>              - Current implementation (baseline)");
            Console.WriteLine("  2. HashSet<T>           - Proposed solution");
            Console.WriteLine("  3. Dictionary<K,V>      - Alternative with IntPtr key");
            Console.WriteLine("  4. ConcurrentDictionary - Lock-free alternative");
            Console.WriteLine();
            Console.WriteLine("Scenarios tested:");
            Console.WriteLine("  A. Sequential Removal   - Worst case: N objects removed one-by-one");
            Console.WriteLine("  B. Bulk Disposal        - Removing from end (current pattern)");
            Console.WriteLine("  C. Mixed Operations     - 70% adds, 30% removes interleaved");
            Console.WriteLine("  D. Memory Allocation    - Memory overhead and GC pressure");
            Console.WriteLine();
            Console.WriteLine("Object counts: 10, 100, 1,000, 10,000");
            Console.WriteLine("Target frameworks: .NET Framework 4.8, .NET 10.0");
            Console.WriteLine();

            // Configure exporters for comprehensive results
            var config = DefaultConfig.Instance
                .AddExporter(MarkdownExporter.GitHub)
                .AddExporter(HtmlExporter.Default)
                .AddExporter(CsvExporter.Default);

            if (args.Length > 0)
            {
                // Run specific benchmark class if provided
                switch (args[0].ToLowerInvariant())
                {
                    case "removal":
                    case "main":
                        Console.WriteLine("Running main removal benchmarks (Scenarios A, B, C)...");
                        BenchmarkRunner.Run<RemoveObjectListBenchmark>(config);
                        break;

                    case "memory":
                    case "mem":
                        Console.WriteLine("Running memory benchmarks (Scenario D)...");
                        BenchmarkRunner.Run<MemoryBenchmark>(config);
                        break;

                    case "all":
                        Console.WriteLine("Running all benchmarks...");
                        BenchmarkRunner.Run<RemoveObjectListBenchmark>(config);
                        BenchmarkRunner.Run<MemoryBenchmark>(config);
                        break;

                    default:
                        Console.WriteLine($"Unknown benchmark: {args[0]}");
                        ShowUsage();
                        return;
                }
            }
            else
            {
                // Default: Run all benchmarks
                Console.WriteLine("Running all benchmarks...");
                Console.WriteLine();
                BenchmarkRunner.Run<RemoveObjectListBenchmark>(config);
                BenchmarkRunner.Run<MemoryBenchmark>(config);
            }

            Console.WriteLine();
            Console.WriteLine("=======================================================");
            Console.WriteLine("Benchmarks completed!");
            Console.WriteLine("Results have been saved to BenchmarkDotNet.Artifacts/");
            Console.WriteLine("=======================================================");
        }

        static void ShowUsage()
        {
            Console.WriteLine("Usage: NetOffice.Benchmarks [benchmark]");
            Console.WriteLine();
            Console.WriteLine("Benchmarks:");
            Console.WriteLine("  removal, main  - Run main removal benchmarks (Scenarios A, B, C)");
            Console.WriteLine("  memory, mem    - Run memory benchmarks (Scenario D)");
            Console.WriteLine("  all            - Run all benchmarks (default)");
            Console.WriteLine();
        }
    }
}

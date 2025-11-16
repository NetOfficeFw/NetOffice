# NetOffice.Benchmarks

Performance benchmarks for NetOffice Core, specifically addressing [Issue #221](https://github.com/NetOfficeFw/NetOffice/issues/221): `Core.RemoveObjectFromList` performance bottleneck.

## Overview

This benchmark project compares the performance of different collection implementations for the `_globalObjectList` in `Core.cs`. The current implementation uses `List<ICOMObject>`, which results in O(n²) complexity when disposing parent objects with multiple children.

## Problem Statement

The current `List<ICOMObject>` implementation causes performance issues because:
- Each `Remove()` operation is O(n)
- Disposing N children results in cumulative O(n²) time complexity
- This becomes a bottleneck when working with large COM object hierarchies

## Implementations Tested

1. **List** (Baseline) - Current implementation
   - Time: O(n) per removal
   - Memory: O(n)
   - Thread-safety: Manual locking

2. **HashSet** (Proposed)
   - Time: O(1) per removal (average)
   - Memory: O(n) with higher constant factor
   - Thread-safety: Manual locking
   - Requires proper `GetHashCode()` and `Equals()` implementation

3. **Dictionary** (Alternative)
   - Time: O(1) per removal by key
   - Memory: O(n)
   - Thread-safety: Manual locking
   - Uses object's hash code as key for guaranteed lookup speed

4. **ConcurrentDictionary** (Lock-free)
   - Time: O(1) per removal (average)
   - Memory: O(n) with higher overhead
   - Thread-safety: Built-in lock-free operations
   - Reduces lock contention in multi-threaded scenarios

## Benchmark Scenarios

### Scenario A: Sequential Removal
Simulates disposing a parent with N children (worst case for List).
- **Test**: Add N objects, then remove them all one by one
- **Expected**: List shows O(n²) complexity, others show O(n)

### Scenario B: Bulk Disposal
Tests `DisposeAllCOMProxies()` behavior with removal from end.
- **Test**: Add N objects, remove from end in while loop
- **Expected**: Better performance for List than sequential, similar for others

### Scenario C: Mixed Operations
Real-world usage pattern with interleaved operations.
- **Test**: 70% adds, 30% removes with random order
- **Expected**: Shows practical performance under lock contention

### Scenario D: Memory Allocation
Compares memory overhead and GC pressure.
- **Test**: Memory footprint and allocation patterns
- **Expected**: HashSet/Dictionary have higher base memory but better scaling

## Test Parameters

- **Object counts**: 10, 100, 1,000, 10,000
- **Target frameworks**: .NET Framework 4.8, .NET 8.0
- **Metrics**:
  - Mean/Median execution time
  - Memory allocations
  - GC collections
  - Operations per second

## Running the Benchmarks

### Prerequisites
- .NET SDK 8.0 or later (for .NET 10 support)
- .NET Framework 4.8 Developer Pack (for .NET Framework targets)
- Admin/elevated permissions may be required for accurate profiling

### Commands

Run all benchmarks:
```bash
cd Tests/NetOffice.Benchmarks
dotnet run -c Release
```

Run specific benchmark suite:
```bash
# Main removal benchmarks (Scenarios A, B, C)
dotnet run -c Release -- removal

# Memory benchmarks (Scenario D)
dotnet run -c Release -- memory

# All benchmarks explicitly
dotnet run -c Release -- all
```

### Important Notes

1. **Always use Release configuration** for accurate benchmarks
2. **Close unnecessary applications** to reduce noise
3. **Run on battery power** (laptops) or disable power-saving features
4. **Disable antivirus scanning** of the output directory if possible
5. Benchmarks will take **10-30 minutes** to complete depending on hardware

## Output

Results are saved to `BenchmarkDotNet.Artifacts/results/`:
- `*.html` - Interactive HTML reports
- `*.md` - GitHub-flavored Markdown reports
- `*.csv` - Raw data for further analysis
- `*-report-github.md` - Summary report for GitHub issues

## Files

- **`Program.cs`** - Entry point and benchmark runner
- **`RemoveObjectListBenchmark.cs`** - Main removal benchmarks (Scenarios A, B, C)
- **`MemoryBenchmark.cs`** - Memory allocation benchmarks (Scenario D)
- **`CoreVariants.cs`** - Different collection implementations
- **`MockCOMObject.cs`** - Minimal ICOMObject implementation for testing

## Expected Results

Based on theoretical analysis:

| Scenario | List (O(n²)) | HashSet/Dict (O(n)) | Improvement |
|----------|-------------|---------------------|-------------|
| 10 objects | ~1 ms | ~0.1 ms | 10x |
| 100 objects | ~10 ms | ~1 ms | 10x |
| 1,000 objects | ~1 s | ~10 ms | 100x |
| 10,000 objects | ~100 s | ~100 ms | 1000x |

## Next Steps

1. Run benchmarks and analyze results
2. Create detailed performance analysis report
3. Update Issue #221 with findings
4. Recommend implementation change based on data
5. Consider implementing the chosen solution

## References

- Issue: https://github.com/NetOfficeFw/NetOffice/issues/221
- Code: `Source/NetOffice/Core.cs:96, 1346-1367`
- BenchmarkDotNet: https://benchmarkdotnet.org/
- Plan: `.github/prompts/plan-NetOfficeBenchmark.md`

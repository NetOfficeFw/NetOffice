using System;
using System.Collections.Generic;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using NetOffice;

namespace NetOffice.Benchmarks
{
    /// <summary>
    /// Scenario D: Memory Allocation Benchmark
    /// Compares memory overhead and GC pressure between different collection implementations.
    /// </summary>
    [MemoryDiagnoser]
    [SimpleJob(RuntimeMoniker.Net48)]
    [SimpleJob(RuntimeMoniker.Net80)]
    [SimpleJob(RuntimeMoniker.Net10_0)]
    public class MemoryBenchmark
    {
        [Params(10, 100, 1000, 10000)]
        public int ObjectCount;

        /// <summary>
        /// Memory footprint for List implementation
        /// </summary>
        [Benchmark(Baseline = true, Description = "List - Memory Footprint")]
        public int MemoryFootprint_List()
        {
            var variant = new ListCoreVariant();
            var objects = new List<ICOMObject>(ObjectCount);

            // Create and add objects
            for (int i = 0; i < ObjectCount; i++)
            {
                var obj = new MockCOMObject();
                objects.Add(obj);
                variant.AddObject(obj);
            }

            // Perform some operations to measure steady-state memory
            int count = variant.Count;

            // Remove half the objects to measure memory behavior
            for (int i = 0; i < ObjectCount / 2; i++)
            {
                variant.RemoveObject(objects[i]);
            }

            return variant.Count;
        }

        /// <summary>
        /// Memory footprint for HashSet implementation
        /// </summary>
        [Benchmark(Description = "HashSet - Memory Footprint")]
        public int MemoryFootprint_HashSet()
        {
            var variant = new HashSetCoreVariant();
            var objects = new List<ICOMObject>(ObjectCount);

            // Create and add objects
            for (int i = 0; i < ObjectCount; i++)
            {
                var obj = new MockCOMObject();
                objects.Add(obj);
                variant.AddObject(obj);
            }

            // Perform some operations to measure steady-state memory
            int count = variant.Count;

            // Remove half the objects to measure memory behavior
            for (int i = 0; i < ObjectCount / 2; i++)
            {
                variant.RemoveObject(objects[i]);
            }

            return variant.Count;
        }

        /// <summary>
        /// Memory footprint for Dictionary implementation
        /// </summary>
        [Benchmark(Description = "Dictionary - Memory Footprint")]
        public int MemoryFootprint_Dictionary()
        {
            var variant = new DictionaryCoreVariant();
            var objects = new List<ICOMObject>(ObjectCount);

            // Create and add objects
            for (int i = 0; i < ObjectCount; i++)
            {
                var obj = new MockCOMObject();
                objects.Add(obj);
                variant.AddObject(obj);
            }

            // Perform some operations to measure steady-state memory
            int count = variant.Count;

            // Remove half the objects to measure memory behavior
            for (int i = 0; i < ObjectCount / 2; i++)
            {
                variant.RemoveObject(objects[i]);
            }

            return variant.Count;
        }

        /// <summary>
        /// Memory footprint for ConcurrentDictionary implementation
        /// </summary>
        [Benchmark(Description = "ConcurrentDictionary - Memory Footprint")]
        public int MemoryFootprint_ConcurrentDictionary()
        {
            var variant = new ConcurrentDictionaryCoreVariant();
            var objects = new List<ICOMObject>(ObjectCount);

            // Create and add objects
            for (int i = 0; i < ObjectCount; i++)
            {
                var obj = new MockCOMObject();
                objects.Add(obj);
                variant.AddObject(obj);
            }

            // Perform some operations to measure steady-state memory
            int count = variant.Count;

            // Remove half the objects to measure memory behavior
            for (int i = 0; i < ObjectCount / 2; i++)
            {
                variant.RemoveObject(objects[i]);
            }

            return variant.Count;
        }

        /// <summary>
        /// Test allocation patterns - List
        /// </summary>
        [Benchmark(Description = "List - Allocation Pattern")]
        public void AllocationPattern_List()
        {
            var variant = new ListCoreVariant();

            // Simulate typical lifecycle: add, use, remove
            for (int cycle = 0; cycle < 10; cycle++)
            {
                var objects = new List<ICOMObject>();

                // Add phase
                for (int i = 0; i < ObjectCount / 10; i++)
                {
                    var obj = new MockCOMObject();
                    objects.Add(obj);
                    variant.AddObject(obj);
                }

                // Remove phase
                foreach (var obj in objects)
                {
                    variant.RemoveObject(obj);
                }
            }
        }

        /// <summary>
        /// Test allocation patterns - HashSet
        /// </summary>
        [Benchmark(Description = "HashSet - Allocation Pattern")]
        public void AllocationPattern_HashSet()
        {
            var variant = new HashSetCoreVariant();

            // Simulate typical lifecycle: add, use, remove
            for (int cycle = 0; cycle < 10; cycle++)
            {
                var objects = new List<ICOMObject>();

                // Add phase
                for (int i = 0; i < ObjectCount / 10; i++)
                {
                    var obj = new MockCOMObject();
                    objects.Add(obj);
                    variant.AddObject(obj);
                }

                // Remove phase
                foreach (var obj in objects)
                {
                    variant.RemoveObject(obj);
                }
            }
        }

        /// <summary>
        /// Test allocation patterns - Dictionary
        /// </summary>
        [Benchmark(Description = "Dictionary - Allocation Pattern")]
        public void AllocationPattern_Dictionary()
        {
            var variant = new DictionaryCoreVariant();

            // Simulate typical lifecycle: add, use, remove
            for (int cycle = 0; cycle < 10; cycle++)
            {
                var objects = new List<ICOMObject>();

                // Add phase
                for (int i = 0; i < ObjectCount / 10; i++)
                {
                    var obj = new MockCOMObject();
                    objects.Add(obj);
                    variant.AddObject(obj);
                }

                // Remove phase
                foreach (var obj in objects)
                {
                    variant.RemoveObject(obj);
                }
            }
        }

        /// <summary>
        /// Test allocation patterns - ConcurrentDictionary
        /// </summary>
        [Benchmark(Description = "ConcurrentDictionary - Allocation Pattern")]
        public void AllocationPattern_ConcurrentDictionary()
        {
            var variant = new ConcurrentDictionaryCoreVariant();

            // Simulate typical lifecycle: add, use, remove
            for (int cycle = 0; cycle < 10; cycle++)
            {
                var objects = new List<ICOMObject>();

                // Add phase
                for (int i = 0; i < ObjectCount / 10; i++)
                {
                    var obj = new MockCOMObject();
                    objects.Add(obj);
                    variant.AddObject(obj);
                }

                // Remove phase
                foreach (var obj in objects)
                {
                    variant.RemoveObject(obj);
                }
            }
        }
    }
}

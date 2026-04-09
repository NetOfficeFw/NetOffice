using System;
using System.Collections.Generic;
using System.Linq;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using NetOffice;

namespace NetOffice.Benchmarks
{
    /// <summary>
    /// Benchmarks for RemoveObjectFromList performance comparing different collection implementations.
    /// Tests various scenarios: Sequential removal, Bulk disposal, Mixed operations, and Memory allocation.
    /// </summary>
    [MemoryDiagnoser]
    [SimpleJob(RuntimeMoniker.Net48)]
    [SimpleJob(RuntimeMoniker.Net80)]
    [SimpleJob(RuntimeMoniker.Net10_0)]
    public class RemoveObjectListBenchmark
    {
        [Params(10, 100, 1000, 10000)]
        public int ObjectCount;

        private List<ICOMObject>? _objects;

        [GlobalSetup]
        public void GlobalSetup()
        {
            // Pre-create objects to avoid allocation overhead in benchmarks
            _objects = new List<ICOMObject>(ObjectCount);
            for (int i = 0; i < ObjectCount; i++)
            {
                _objects.Add(new MockCOMObject());
            }
        }

        #region Scenario A: Sequential Removal (Worst case for List)

        /// <summary>
        /// Scenario A: Sequential Removal with List (Baseline - Current Implementation)
        /// Simulates disposing a parent with N children.
        /// Expected: O(n²) complexity - performance degrades quadratically
        /// </summary>
        [Benchmark(Baseline = true, Description = "List - Sequential Removal")]
        public void SequentialRemoval_List()
        {
            var variant = new ListCoreVariant();

            // Add all objects
            foreach (var obj in _objects!)
            {
                variant.AddObject(obj);
            }

            // Remove all objects one by one (simulates child disposal)
            foreach (var obj in _objects)
            {
                variant.RemoveObject(obj);
            }
        }

        /// <summary>
        /// Scenario A: Sequential Removal with HashSet
        /// Expected: O(n) complexity - linear performance
        /// </summary>
        [Benchmark(Description = "HashSet - Sequential Removal")]
        public void SequentialRemoval_HashSet()
        {
            var variant = new HashSetCoreVariant();

            // Add all objects
            foreach (var obj in _objects!)
            {
                variant.AddObject(obj);
            }

            // Remove all objects one by one
            foreach (var obj in _objects)
            {
                variant.RemoveObject(obj);
            }
        }

        /// <summary>
        /// Scenario A: Sequential Removal with Dictionary
        /// Expected: O(n) complexity - linear performance
        /// </summary>
        [Benchmark(Description = "Dictionary - Sequential Removal")]
        public void SequentialRemoval_Dictionary()
        {
            var variant = new DictionaryCoreVariant();

            // Add all objects
            foreach (var obj in _objects!)
            {
                variant.AddObject(obj);
            }

            // Remove all objects one by one
            foreach (var obj in _objects)
            {
                variant.RemoveObject(obj);
            }
        }

        /// <summary>
        /// Scenario A: Sequential Removal with ConcurrentDictionary
        /// Expected: O(n) complexity - linear performance with reduced lock contention
        /// </summary>
        [Benchmark(Description = "ConcurrentDictionary - Sequential Removal")]
        public void SequentialRemoval_ConcurrentDictionary()
        {
            var variant = new ConcurrentDictionaryCoreVariant();

            // Add all objects
            foreach (var obj in _objects!)
            {
                variant.AddObject(obj);
            }

            // Remove all objects one by one
            foreach (var obj in _objects)
            {
                variant.RemoveObject(obj);
            }
        }

        #endregion

        #region Scenario B: Bulk Disposal

        /// <summary>
        /// Scenario B: Bulk Disposal with List
        /// Tests DisposeAllCOMProxies() pattern - removing from end in while loop
        /// </summary>
        [Benchmark(Description = "List - Bulk Disposal")]
        public void BulkDisposal_List()
        {
            var variant = new ListCoreVariant();

            // Add all objects
            foreach (var obj in _objects!)
            {
                variant.AddObject(obj);
            }

            // Remove from end (better for List than from beginning)
            for (int i = _objects.Count - 1; i >= 0; i--)
            {
                variant.RemoveObject(_objects[i]);
            }
        }

        /// <summary>
        /// Scenario B: Bulk Disposal with HashSet
        /// </summary>
        [Benchmark(Description = "HashSet - Bulk Disposal")]
        public void BulkDisposal_HashSet()
        {
            var variant = new HashSetCoreVariant();

            // Add all objects
            foreach (var obj in _objects!)
            {
                variant.AddObject(obj);
            }

            // Remove from end
            for (int i = _objects.Count - 1; i >= 0; i--)
            {
                variant.RemoveObject(_objects[i]);
            }
        }

        /// <summary>
        /// Scenario B: Bulk Disposal with Dictionary
        /// </summary>
        [Benchmark(Description = "Dictionary - Bulk Disposal")]
        public void BulkDisposal_Dictionary()
        {
            var variant = new DictionaryCoreVariant();

            // Add all objects
            foreach (var obj in _objects!)
            {
                variant.AddObject(obj);
            }

            // Remove from end
            for (int i = _objects.Count - 1; i >= 0; i--)
            {
                variant.RemoveObject(_objects[i]);
            }
        }

        /// <summary>
        /// Scenario B: Bulk Disposal with ConcurrentDictionary
        /// </summary>
        [Benchmark(Description = "ConcurrentDictionary - Bulk Disposal")]
        public void BulkDisposal_ConcurrentDictionary()
        {
            var variant = new ConcurrentDictionaryCoreVariant();

            // Add all objects
            foreach (var obj in _objects!)
            {
                variant.AddObject(obj);
            }

            // Remove from end
            for (int i = _objects.Count - 1; i >= 0; i--)
            {
                variant.RemoveObject(_objects[i]);
            }
        }

        #endregion

        #region Scenario C: Mixed Operations

        /// <summary>
        /// Scenario C: Mixed Operations with List
        /// Real-world usage: 70% adds, 30% removes interleaved
        /// </summary>
        [Benchmark(Description = "List - Mixed Operations")]
        public void MixedOperations_List()
        {
            var variant = new ListCoreVariant();
            var random = new Random(42); // Fixed seed for reproducibility
            var addedObjects = new List<ICOMObject>();

            for (int i = 0; i < ObjectCount; i++)
            {
                if (random.NextDouble() < 0.7 || addedObjects.Count == 0)
                {
                    // Add operation (70% of time)
                    var obj = _objects![i % _objects.Count];
                    variant.AddObject(obj);
                    addedObjects.Add(obj);
                }
                else
                {
                    // Remove operation (30% of time)
                    int removeIndex = random.Next(addedObjects.Count);
                    var objToRemove = addedObjects[removeIndex];
                    variant.RemoveObject(objToRemove);
                    addedObjects.RemoveAt(removeIndex);
                }
            }
        }

        /// <summary>
        /// Scenario C: Mixed Operations with HashSet
        /// </summary>
        [Benchmark(Description = "HashSet - Mixed Operations")]
        public void MixedOperations_HashSet()
        {
            var variant = new HashSetCoreVariant();
            var random = new Random(42);
            var addedObjects = new List<ICOMObject>();

            for (int i = 0; i < ObjectCount; i++)
            {
                if (random.NextDouble() < 0.7 || addedObjects.Count == 0)
                {
                    var obj = _objects![i % _objects.Count];
                    variant.AddObject(obj);
                    if (!addedObjects.Contains(obj))
                        addedObjects.Add(obj);
                }
                else
                {
                    int removeIndex = random.Next(addedObjects.Count);
                    var objToRemove = addedObjects[removeIndex];
                    variant.RemoveObject(objToRemove);
                    addedObjects.RemoveAt(removeIndex);
                }
            }
        }

        /// <summary>
        /// Scenario C: Mixed Operations with Dictionary
        /// </summary>
        [Benchmark(Description = "Dictionary - Mixed Operations")]
        public void MixedOperations_Dictionary()
        {
            var variant = new DictionaryCoreVariant();
            var random = new Random(42);
            var addedObjects = new List<ICOMObject>();

            for (int i = 0; i < ObjectCount; i++)
            {
                if (random.NextDouble() < 0.7 || addedObjects.Count == 0)
                {
                    var obj = _objects![i % _objects.Count];
                    variant.AddObject(obj);
                    if (!addedObjects.Contains(obj))
                        addedObjects.Add(obj);
                }
                else
                {
                    int removeIndex = random.Next(addedObjects.Count);
                    var objToRemove = addedObjects[removeIndex];
                    variant.RemoveObject(objToRemove);
                    addedObjects.RemoveAt(removeIndex);
                }
            }
        }

        /// <summary>
        /// Scenario C: Mixed Operations with ConcurrentDictionary
        /// </summary>
        [Benchmark(Description = "ConcurrentDictionary - Mixed Operations")]
        public void MixedOperations_ConcurrentDictionary()
        {
            var variant = new ConcurrentDictionaryCoreVariant();
            var random = new Random(42);
            var addedObjects = new List<ICOMObject>();

            for (int i = 0; i < ObjectCount; i++)
            {
                if (random.NextDouble() < 0.7 || addedObjects.Count == 0)
                {
                    var obj = _objects![i % _objects.Count];
                    variant.AddObject(obj);
                    if (!addedObjects.Contains(obj))
                        addedObjects.Add(obj);
                }
                else
                {
                    int removeIndex = random.Next(addedObjects.Count);
                    var objToRemove = addedObjects[removeIndex];
                    variant.RemoveObject(objToRemove);
                    addedObjects.RemoveAt(removeIndex);
                }
            }
        }

        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using NetOffice;

namespace NetOffice.Benchmarks
{
    /// <summary>
    /// Base interface for Core collection variants to enable polymorphic benchmarking
    /// </summary>
    internal interface ICoreCollectionVariant
    {
        void AddObject(ICOMObject obj);
        bool RemoveObject(ICOMObject obj);
        void Clear();
        int Count { get; }
    }

    /// <summary>
    /// Variant 1: Current implementation using List (baseline)
    /// Time: O(n) per removal
    /// Memory: O(n)
    /// </summary>
    internal class ListCoreVariant : ICoreCollectionVariant
    {
        private readonly List<ICOMObject> _globalObjectList = new List<ICOMObject>();
        private readonly object _lock = new object();

        public void AddObject(ICOMObject obj)
        {
            lock (_lock)
            {
                _globalObjectList.Add(obj);
            }
        }

        public bool RemoveObject(ICOMObject obj)
        {
            lock (_lock)
            {
                return _globalObjectList.Remove(obj);
            }
        }

        public void Clear()
        {
            lock (_lock)
            {
                _globalObjectList.Clear();
            }
        }

        public int Count
        {
            get
            {
                lock (_lock)
                {
                    return _globalObjectList.Count;
                }
            }
        }
    }

    /// <summary>
    /// Variant 2: HashSet implementation (proposed solution)
    /// Time: O(1) per removal (average)
    /// Memory: O(n) with higher constant factor
    /// Requires proper GetHashCode() and Equals() implementation
    /// </summary>
    internal class HashSetCoreVariant : ICoreCollectionVariant
    {
        private readonly HashSet<ICOMObject> _globalObjectList = new HashSet<ICOMObject>();
        private readonly object _lock = new object();

        public void AddObject(ICOMObject obj)
        {
            lock (_lock)
            {
                _globalObjectList.Add(obj);
            }
        }

        public bool RemoveObject(ICOMObject obj)
        {
            lock (_lock)
            {
                return _globalObjectList.Remove(obj);
            }
        }

        public void Clear()
        {
            lock (_lock)
            {
                _globalObjectList.Clear();
            }
        }

        public int Count
        {
            get
            {
                lock (_lock)
                {
                    return _globalObjectList.Count;
                }
            }
        }
    }

    /// <summary>
    /// Variant 3: Dictionary keyed by IntPtr (alternative)
    /// Time: O(1) per removal by key
    /// Memory: O(n)
    /// Benefit: Can key by COM pointer for guaranteed uniqueness
    /// </summary>
    internal class DictionaryCoreVariant : ICoreCollectionVariant
    {
        private readonly Dictionary<int, ICOMObject> _globalObjectList = new Dictionary<int, ICOMObject>();
        private readonly object _lock = new object();

        public void AddObject(ICOMObject obj)
        {
            lock (_lock)
            {
                int key = obj.GetHashCode();
                _globalObjectList[key] = obj;
            }
        }

        public bool RemoveObject(ICOMObject obj)
        {
            lock (_lock)
            {
                int key = obj.GetHashCode();
                return _globalObjectList.Remove(key);
            }
        }

        public void Clear()
        {
            lock (_lock)
            {
                _globalObjectList.Clear();
            }
        }

        public int Count
        {
            get
            {
                lock (_lock)
                {
                    return _globalObjectList.Count;
                }
            }
        }
    }

    /// <summary>
    /// Variant 4: ConcurrentDictionary (lock-free alternative)
    /// Time: O(1) per removal (average)
    /// Memory: O(n) with higher overhead
    /// Benefit: Reduces lock contention with built-in thread-safety
    /// </summary>
    internal class ConcurrentDictionaryCoreVariant : ICoreCollectionVariant
    {
        private readonly ConcurrentDictionary<int, ICOMObject> _globalObjectList =
            new ConcurrentDictionary<int, ICOMObject>();

        public void AddObject(ICOMObject obj)
        {
            int key = obj.GetHashCode();
            _globalObjectList[key] = obj;
        }

        public bool RemoveObject(ICOMObject obj)
        {
            int key = obj.GetHashCode();
            return _globalObjectList.TryRemove(key, out _);
        }

        public void Clear()
        {
            _globalObjectList.Clear();
        }

        public int Count => _globalObjectList.Count;
    }
}

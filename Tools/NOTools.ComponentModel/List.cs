using System;
using System.ComponentModel;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime;
using System.Security;
using System.Threading;

namespace NOTools.ComponentModel
{
    [DebuggerDisplay("Count = {Count}")]
    [Serializable]
    public class List<T> : IList<T>, ICollection<T>, IEnumerable<T>, IList, ICollection, IEnumerable
    {
        #region Embedded Types

        [Serializable]
        internal class SynchronizedList : IList<T>, ICollection<T>, IEnumerable<T>, IEnumerable
        {
            private List<T> _list;
            private object _root;
            public int Count
            {
                get
                {
                    int count;
                    lock (this._root)
                    {
                        count = this._list.Count;
                    }
                    return count;
                }
            }
            public bool IsReadOnly
            {
                get
                {
                    return ((ICollection<T>)this._list).IsReadOnly;
                }
            }
            public T this[int index]
            {
                get
                {
                    T result;
                    lock (this._root)
                    {
                        result = this._list[index];
                    }
                    return result;
                }
                set
                {
                    lock (this._root)
                    {
                        this._list[index] = value;
                    }
                }
            }
            internal SynchronizedList(List<T> list)
            {
                this._list = list;
                this._root = ((ICollection)list).SyncRoot;
            }
            public void Add(T item)
            {
                lock (this._root)
                {
                    this._list.Add(item);
                }
            }
            public void Clear()
            {
                lock (this._root)
                {
                    this._list.Clear();
                }
            }
            public bool Contains(T item)
            {
                bool result;
                lock (this._root)
                {
                    result = this._list.Contains(item);
                }
                return result;
            }
            public void CopyTo(T[] array, int arrayIndex)
            {
                lock (this._root)
                {
                    this._list.CopyTo(array, arrayIndex);
                }
            }
            public bool Remove(T item)
            {
                bool result;
                lock (this._root)
                {
                    result = this._list.Remove(item);
                }
                return result;
            }
            IEnumerator IEnumerable.GetEnumerator()
            {
                IEnumerator result;
                lock (this._root)
                {
                    result = this._list.GetEnumerator();
                }
                return result;
            }
            IEnumerator<T> IEnumerable<T>.GetEnumerator()
            {
                IEnumerator<T> enumerator;
                lock (this._root)
                {
                    enumerator = ((IEnumerable<T>)this._list).GetEnumerator();
                }
                return enumerator;
            }
            public int IndexOf(T item)
            {
                int result;
                lock (this._root)
                {
                    result = this._list.IndexOf(item);
                }
                return result;
            }
            public void Insert(int index, T item)
            {
                lock (this._root)
                {
                    this._list.Insert(index, item);
                }
            }
            public void RemoveAt(int index)
            {
                lock (this._root)
                {
                    this._list.RemoveAt(index);
                }
            }
        }
    
        [Serializable]
        public struct Enumerator : IEnumerator<T>, IDisposable, IEnumerator
        {
            private List<T> list;
            private int index;
            private int version;
            private T current;
            public T Current
            {
                [TargetedPatchingOptOut("Performance critical to inline this type of method across NGen image boundaries")]
                get
                {
                    return this.current;
                }
            }
            object IEnumerator.Current
            {
                get
                {
                    if (this.index == 0 || this.index == this.list._size + 1)
                        throw new InvalidOperationException();
                    
                    return this.Current;
                }
            }
            internal Enumerator(List<T> list)
            {
                this.list = list;
                this.index = 0;
                this.version = list._version;
                this.current = default(T);
            }
            [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
            public void Dispose()
            {
            }
            public bool MoveNext()
            {
                List<T> list = this.list;
                if (this.version == list._version && this.index < list._size)
                {
                    this.current = list._items[this.index];
                    this.index++;
                    return true;
                }
                return this.MoveNextRare();
            }
            private bool MoveNextRare()
            {
                if (this.version != this.list._version)
                {
                     throw new InvalidOperationException();
                }
                this.index = this.list._size + 1;
                this.current = default(T);
                return false;
            }
            void IEnumerator.Reset()
            {
                if (this.version != this.list._version)
                {
                     throw new InvalidOperationException();
                }
                this.index = 0;
                this.current = default(T);
            }
        }

        #endregion

        #region Fields

        private T[] _items;
        private int _size;
        private int _version;
        [NonSerialized]
        private object _syncRoot;
        private static readonly T[] _emptyArray = new T[0];
        private const int _defaultCapacity = 4;

        #endregion

        #region Ctor

        public List()
        {
            this._items = List<T>._emptyArray;
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        public List(int capacity)
        {
            if (capacity < 0)
            {
                throw new InvalidOperationException();
            }
            this._items = new T[capacity];
        }

        public List(IEnumerable<T> collection)
        {
            if (collection == null)
            {
                throw new InvalidOperationException();
            }
            ICollection<T> collection2 = collection as ICollection<T>;
            if (collection2 != null)
            {
                int count = collection2.Count;
                this._items = new T[count];
                collection2.CopyTo(this._items, 0);
                this._size = count;
                return;
            }
            this._size = 0;
            this._items = new T[4];
            using (IEnumerator<T> enumerator = collection.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    this.Add(enumerator.Current);
                }
            }
        }

        #endregion

        #region Properties

        public int Capacity
        {
            [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
            get
            {
                return this._items.Length;
            }
            set
            {
                if (value < this._size)
                {
                    throw new InvalidOperationException();
                }
                if (value != this._items.Length)
                {
                    if (value > 0)
                    {
                        T[] array = new T[value];
                        if (this._size > 0)
                        {
                            Array.Copy(this._items, 0, array, 0, this._size);
                        }
                        this._items = array;
                        return;
                    }
                    this._items = List<T>._emptyArray;
                }
            }
        }

        public int Count
        {
            [TargetedPatchingOptOut("Performance critical to inline this type of method across NGen image boundaries")]
            get
            {
                return this._size;
            }
        }

        public T this[int index]
        {
            [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
            get
            {
                if (index >= this._size)
                {
                    throw new InvalidOperationException();
                }
                return this._items[index];
            }
            [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
            set
            {
                if (index >= this._size)
                {
                    throw new InvalidOperationException();
                }
                this._items[index] = value;
                this._version++;
            }
        }

        #endregion

        #region IList

        bool IList.IsFixedSize
        {
            get
            {
                return false;
            }
        }

        bool IList.IsReadOnly
        {
            [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
            get
            {
                return false;
            }
        }

        object IList.this[int index]
        {
            get
            {
                return this[index];
            }
            set
            {
                IfNullAndNullsAreIllegalThenThrow<T>(value);
                try
                {
                    this[index] = (T)((object)value);
                }
                catch (InvalidCastException)
                {
                    throw new InvalidOperationException();
                }
            }
        }

        int IList.Add(object item)
        {
            try
            {
                this.Add((T)((object)item));
            }
            catch (InvalidCastException)
            {
                throw new InvalidOperationException();
            }
            return this.Count - 1;
        }

        [SecuritySafeCritical]
        bool IList.Contains(object item)
        {
            return List<T>.IsCompatibleObject(item) && this.Contains((T)((object)item));
        }
         
        void IList.Insert(int index, object item)
        {
            try
            {
                this.Insert(index, (T)((object)item));
            }
            catch (InvalidCastException)
            {
                throw new InvalidOperationException();
            }
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries"), SecuritySafeCritical]
        int IList.IndexOf(object item)
        {
            if (List<T>.IsCompatibleObject(item))
            {
                return this.IndexOf((T)((object)item));
            }
            return -1;
        }

        internal static IList<T> Synchronized(List<T> list)
        {
            return new List<T>.SynchronizedList(list);
        }

        [SecuritySafeCritical]
        void IList.Remove(object item)
        {
            if (List<T>.IsCompatibleObject(item))
            {
                this.Remove((T)((object)item));
            }
        }

        #endregion

        #region ICollection

        bool ICollection<T>.IsReadOnly
        {
            [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
            get
            {
                return false;
            }
        }

        bool ICollection.IsSynchronized
        {
            get
            {
                return false;
            }
        }

        object ICollection.SyncRoot
        {
            get
            {
                if (this._syncRoot == null)
                {
                    Interlocked.CompareExchange<object>(ref this._syncRoot, new object(), null);
                }
                return this._syncRoot;
            }
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        void ICollection.CopyTo(Array array, int arrayIndex)
        {
            if (array != null && array.Rank != 1)
            {
                throw new InvalidOperationException();
            }
            try
            {
                Array.Copy(this._items, 0, array, arrayIndex, this._size);
            }
            catch (ArrayTypeMismatchException)
            {
                throw new InvalidOperationException();
            }
        }

        #endregion

        #region IEnumerable

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return new List<T>.Enumerator(this);
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        IEnumerator IEnumerable.GetEnumerator()
        {
            return new List<T>.Enumerator(this);
        }

        #endregion

        #region Methods

        public List<TOutput> ConvertAll<TOutput>(Converter<T, TOutput> converter)
        {
            if (converter == null)
            {
                throw new InvalidOperationException();
            }
            List<TOutput> list = new List<TOutput>(this._size);
            for (int i = 0; i < this._size; i++)
            {
                list._items[i] = converter(this._items[i]);
            }
            list._size = this._size;
            return list;
        }
       
        public void Add(T item)
        {
            if (this._size == this._items.Length)
            {
                this.EnsureCapacity(this._size + 1);
            }
            this._items[this._size++] = item;
            this._version++;
        }

        public void AddRange(IEnumerable<T> collection)
        {
            this.InsertRange(this._size, collection);
        }

        public ReadOnlyCollection<T> AsReadOnly()
        {
            return new ReadOnlyCollection<T>(this);
        }

        public int BinarySearch(int index, int count, T item, IComparer<T> comparer)
        {
            if (index < 0)
            {
                 throw new InvalidOperationException();
            }
            if (count < 0)
            {
                 throw new InvalidOperationException();
            }
            if (this._size - index < count)
            {
                throw new InvalidOperationException();
            }
            return Array.BinarySearch<T>(this._items, index, count, item, comparer);
        }

        public int BinarySearch(T item)
        {
            return this.BinarySearch(0, this.Count, item, null);
        }

        public int BinarySearch(T item, IComparer<T> comparer)
        {
            return this.BinarySearch(0, this.Count, item, comparer);
        }

        public void Clear()
        {
            if (this._size > 0)
            {
                Array.Clear(this._items, 0, this._size);
                this._size = 0;
            }
            this._version++;
        }
  
        public bool Contains(T item)
        {
            if (item == null)
            {
                for (int i = 0; i < this._size; i++)
                {
                    if (this._items[i] == null)
                    {
                        return true;
                    }
                }
                return false;
            }
            EqualityComparer<T> @default = EqualityComparer<T>.Default;
            for (int j = 0; j < this._size; j++)
            {
                if (@default.Equals(this._items[j], item))
                {
                    return true;
                }
            }
            return false;
        }
      
        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        public void CopyTo(T[] array)
        {
            this.CopyTo(array, 0);
        }
      
        public void CopyTo(int index, T[] array, int arrayIndex, int count)
        {
            if (this._size - index < count)
            {
                throw new InvalidOperationException();
            }
            Array.Copy(this._items, index, array, arrayIndex, count);
        }

        public void CopyTo(T[] array, int arrayIndex)
        {
            Array.Copy(this._items, 0, array, arrayIndex, this._size);
        }

        private void EnsureCapacity(int min)
        {
            if (this._items.Length < min)
            {
                int num = (this._items.Length == 0) ? 4 : (this._items.Length * 2);
                if (num < min)
                {
                    num = min;
                }
                this.Capacity = num;
            }
        }

        public bool Exists(Predicate<T> match)
        {
            return this.FindIndex(match) != -1;
        }

        public T Find(Predicate<T> match)
        {
            if (match == null)
            {
                throw new InvalidOperationException();
            }
            for (int i = 0; i < this._size; i++)
            {
                if (match(this._items[i]))
                {
                    return this._items[i];
                }
            }
            return default(T);
        }

        public List<T> FindAll(Predicate<T> match)
        {
            if (match == null)
            {
                 throw new InvalidOperationException();
            }
            List<T> list = new List<T>();
            for (int i = 0; i < this._size; i++)
            {
                if (match(this._items[i]))
                {
                    list.Add(this._items[i]);
                }
            }
            return list;
        }

        public int FindIndex(Predicate<T> match)
        {
            return this.FindIndex(0, this._size, match);
        }

        public int FindIndex(int startIndex, Predicate<T> match)
        {
            return this.FindIndex(startIndex, this._size - startIndex, match);
        }

        public int FindIndex(int startIndex, int count, Predicate<T> match)
        {
            if (startIndex > this._size)
            {
               throw new InvalidOperationException();
            }
            if (count < 0 || startIndex > this._size - count)
            {
                throw new InvalidOperationException();
            }
            if (match == null)
            {  
              throw new InvalidOperationException();
            }
            int num = startIndex + count;
            for (int i = startIndex; i < num; i++)
            {
                if (match(this._items[i]))
                {
                    return i;
                }
            }
            return -1;
        }

        public T FindLast(Predicate<T> match)
        {
            if (match == null)
            {
                throw new InvalidOperationException();
            }
            for (int i = this._size - 1; i >= 0; i--)
            {
                if (match(this._items[i]))
                {
                    return this._items[i];
                }
            }
            return default(T);
        }

        public int FindLastIndex(Predicate<T> match)
        {
            return this.FindLastIndex(this._size - 1, this._size, match);
        }

        public int FindLastIndex(int startIndex, Predicate<T> match)
        {
            return this.FindLastIndex(startIndex, startIndex + 1, match);
        }

        public int FindLastIndex(int startIndex, int count, Predicate<T> match)
        {
            if (match == null)
            {
               throw new InvalidOperationException();
            }
            if (this._size == 0)
            {
                if (startIndex != -1)
                { 
                    throw new InvalidOperationException();
                }
            }
            else
            {
                if (startIndex >= this._size)
                {
                    throw new InvalidOperationException();
                }
            }
            if (count < 0 || startIndex - count + 1 < 0)
            {
                throw new InvalidOperationException();
            }
            int num = startIndex - count;
            for (int i = startIndex; i > num; i--)
            {
                if (match(this._items[i]))
                {
                    return i;
                }
            }
            return -1;
        }

        public void ForEach(Action<T> action)
        {
            if (action == null)
            {
                 throw new InvalidOperationException();
            }
            for (int i = 0; i < this._size; i++)
            {
                action(this._items[i]);
            }
        }
       
        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        public List<T>.Enumerator GetEnumerator()
        {
            return new List<T>.Enumerator(this);
        }

        public List<T> GetRange(int index, int count)
        {
            if (index < 0)
            {
                 throw new InvalidOperationException();
            }
            if (count < 0)
            {
                throw new InvalidOperationException();
            }
            if (this._size - index < count)
            {
                throw new InvalidOperationException();
            }
            List<T> list = new List<T>(count);
            Array.Copy(this._items, index, list._items, 0, count);
            list._size = count;
            return list;
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        public int IndexOf(T item)
        {
            return Array.IndexOf<T>(this._items, item, 0, this._size);
        }
         
        public int IndexOf(T item, int index)
        {
            if (index > this._size)
            {
                throw new InvalidOperationException();
            }
            return Array.IndexOf<T>(this._items, item, index, this._size - index);
        }

        public int IndexOf(T item, int index, int count)
        {
            if (index > this._size)
            {
                 throw new InvalidOperationException();
            }
            if (count < 0 || index > this._size - count)
            {
                 throw new InvalidOperationException();
            }
            return Array.IndexOf<T>(this._items, item, index, count);
        }

        public void Insert(int index, T item)
        {
            if (index > this._size)
            {
                 throw new InvalidOperationException();
            }
            if (this._size == this._items.Length)
            {
                this.EnsureCapacity(this._size + 1);
            }
            if (index < this._size)
            {
                Array.Copy(this._items, index, this._items, index + 1, this._size - index);
            }
            this._items[index] = item;
            this._size++;
            this._version++;
        }
        
        public void InsertRange(int index, IEnumerable<T> collection)
        {
            if (collection == null)
            {
                throw new InvalidOperationException();
            }
            if (index > this._size)
            {
                 throw new InvalidOperationException();
            }
            ICollection<T> collection2 = collection as ICollection<T>;
            if (collection2 != null)
            {
                int count = collection2.Count;
                if (count > 0)
                {
                    this.EnsureCapacity(this._size + count);
                    if (index < this._size)
                    {
                        Array.Copy(this._items, index, this._items, index + count, this._size - index);
                    }
                    if (this == collection2)
                    {
                        Array.Copy(this._items, 0, this._items, index, index);
                        Array.Copy(this._items, index + count, this._items, index * 2, this._size - index);
                    }
                    else
                    {
                        T[] array = new T[count];
                        collection2.CopyTo(array, 0);
                        array.CopyTo(this._items, index);
                    }
                    this._size += count;
                }
            }
            else
            {
                using (IEnumerator<T> enumerator = collection.GetEnumerator())
                {
                    while (enumerator.MoveNext())
                    {
                        this.Insert(index++, enumerator.Current);
                    }
                }
            }
            this._version++;
        }

        public int LastIndexOf(T item)
        {
            if (this._size == 0)
            {
                return -1;
            }
            return this.LastIndexOf(item, this._size - 1, this._size);
        }

        public int LastIndexOf(T item, int index)
        {
            if (index >= this._size)
            {
                throw new InvalidOperationException();
            }
            return this.LastIndexOf(item, index, index + 1);
        }

        public int LastIndexOf(T item, int index, int count)
        {
            if (this.Count != 0 && index < 0)
            {
                 throw new InvalidOperationException();
            }
            if (this.Count != 0 && count < 0)
            {
                throw new InvalidOperationException();
            }
            if (this._size == 0)
            {
                return -1;
            }
            if (index >= this._size)
            {
                throw new InvalidOperationException();
            }
            if (count > index + 1)
            {
                 throw new InvalidOperationException();
            }
            return Array.LastIndexOf<T>(this._items, item, index, count);
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        public bool Remove(T item)
        {
            int num = this.IndexOf(item);
            if (num >= 0)
            {
                this.RemoveAt(num);
                return true;
            }
            return false;
        }

        public int RemoveAll(Predicate<T> match)
        {
            if (match == null)
            {
                throw new InvalidOperationException();
            }
            int num = 0;
            while (num < this._size && !match(this._items[num]))
            {
                num++;
            }
            if (num >= this._size)
            {
                return 0;
            }
            int i = num + 1;
            while (i < this._size)
            {
                while (i < this._size && match(this._items[i]))
                {
                    i++;
                }
                if (i < this._size)
                {
                    this._items[num++] = this._items[i++];
                }
            }
            Array.Clear(this._items, num, this._size - num);
            int result = this._size - num;
            this._size = num;
            this._version++;
            return result;
        }

        public void RemoveAt(int index)
        {
            if (index >= this._size)
            {
                throw new InvalidOperationException();
            }
            this._size--;
            if (index < this._size)
            {
                Array.Copy(this._items, index + 1, this._items, index, this._size - index);
            }
            this._items[this._size] = default(T);
            this._version++;
        }

        public void RemoveRange(int index, int count)
        {
            if (index < 0)
            {
                 throw new InvalidOperationException();
            }
            if (count < 0)
            {
                 throw new InvalidOperationException();
            }
            if (this._size - index < count)
            {
                 throw new InvalidOperationException();
            }
            if (count > 0)
            {
                this._size -= count;
                if (index < this._size)
                {
                    Array.Copy(this._items, index + count, this._items, index, this._size - index);
                }
                Array.Clear(this._items, this._size, count);
                this._version++;
            }
        }

        public void Reverse()
        {
            this.Reverse(0, this.Count);
        }

        public void Reverse(int index, int count)
        {
            if (index < 0)
            {
                throw new InvalidOperationException();
            }
            if (count < 0)
            {
                 throw new InvalidOperationException();
            }
            if (this._size - index < count)
            {
                 throw new InvalidOperationException();
            }
            Array.Reverse(this._items, index, count);
            this._version++;
        }

        public void Sort()
        {
            this.Sort(0, this.Count, null);
        }

        public void Sort(IComparer<T> comparer)
        {
            this.Sort(0, this.Count, comparer);
        }

        public void Sort(int index, int count, IComparer<T> comparer)
        {
            if (index < 0)
            {
               throw new InvalidOperationException();
            }
            if (count < 0)
            {
                throw new InvalidOperationException();
            }
            if (this._size - index < count)
            {
                throw new InvalidOperationException();
            }
            Array.Sort<T>(this._items, index, count, comparer);
            this._version++;
        }

        public void Sort(Comparison<T> comparison)
        {
            if (comparison == null)
                throw new InvalidOperationException();
            
            if (this._size > 0)
            {
                IComparer<T> comparer = new InternalArray.FunctorComparer<T>(comparison);
                Array.Sort<T>(this._items, 0, this._size, comparer);
            }
        }

        public T[] ToArray()
        {
            T[] array = new T[this._size];
            Array.Copy(this._items, 0, array, 0, this._size);
            return array;
        }

        public void TrimExcess()
        {
            int num = (int)((double)this._items.Length * 0.9);
            if (this._size < num)
            {
                this.Capacity = this._size;
            }
        }

        public bool TrueForAll(Predicate<T> match)
        {
            if (match == null)
            {
                 throw new InvalidOperationException();
            }
            for (int i = 0; i < this._size; i++)
            {
                if (!match(this._items[i]))
                {
                    return false;
                }
            }
            return true;
        }

        #endregion

        #region Static Methods

        private static void IfNullAndNullsAreIllegalThenThrow<TT>(object value)
        {
            if (value == null && default(TT) != null)
                throw new NotSupportedException();
        }

        private static bool IsCompatibleObject(object value)
        {
            return value is T || (value == null && default(T) == null);
        }

        #endregion
    }
}

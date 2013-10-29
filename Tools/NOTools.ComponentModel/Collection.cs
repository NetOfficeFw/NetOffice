using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Threading;

namespace NOTools.ComponentModel
{
    [DebuggerDisplay("Count = {Count}"), ComVisible(false)]
    [Serializable]
    public class Collection<T> : IList<T>, ICollection<T>, IEnumerable<T>, IList, ICollection, IEnumerable
    {
        #region Fields

        private IList<T> _items;

        [NonSerialized]
        private object _syncRoot;

        #endregion

        #region Ctor

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        public Collection()
        {
            _items = new List<T>();
        }

        public Collection(IList<T> list)
        {
            if (list == null)
                throw new NotSupportedException();
            _items = list;
        }

        #endregion

        #region Properties

        public int Count
        {
            [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
            get
            {
                return _items.Count;
            }
        }

        public T this[int index]
        {
            [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
            get
            {
                return _items[index];
            }
            set
            {
                if (_items.IsReadOnly)
                {
                    throw new NotSupportedException();
                }
                if (index < 0 || index >= _items.Count)
                {
                    throw new NotSupportedException();
                }
                this.SetItem(index, value);
            }
        }

        #endregion

        #region ICollection

        bool ICollection<T>.IsReadOnly
        {
            get
            {
                return _items.IsReadOnly;
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
                    ICollection collection = _items as ICollection;
                    if (collection != null)
                    {
                        this._syncRoot = collection.SyncRoot;
                    }
                    else
                    {
                        Interlocked.CompareExchange<object>(ref this._syncRoot, new object(), null);
                    }
                }
                return this._syncRoot;
            }
        }

        void ICollection.CopyTo(Array array, int index)
        {
            if (array == null)
                throw new NotSupportedException();
            if (array.Rank != 1)
                throw new NotSupportedException();
            if (array.GetLowerBound(0) != 0)
                throw new NotSupportedException();
            if (index < 0)
                throw new NotSupportedException();
            if (array.Length - index < this.Count)
                throw new NotSupportedException();

            T[] array2 = array as T[];
            if (array2 != null)
            {
                _items.CopyTo(array2, index);
                return;
            }

            Type elementType = array.GetType().GetElementType();
            Type typeFromHandle = typeof(T);
            if (!elementType.IsAssignableFrom(typeFromHandle) && !typeFromHandle.IsAssignableFrom(elementType))
                throw new NotSupportedException();

            object[] array3 = array as object[];
            if (array3 == null)
                throw new NotSupportedException();

            int count = _items.Count;
            try
            {
                for (int i = 0; i < count; i++)
                    array3[index++] = _items[i];
            }
            catch (ArrayTypeMismatchException)
            {
                throw new NotSupportedException();
            }
        }

        #endregion

        #region IList

        protected IList<T> Items
        {
            [TargetedPatchingOptOut("Performance critical to inline this type of method across NGen image boundaries")]
            get
            {
                return _items;
            }
        }


        object IList.this[int index]
        {
            get
            {
                return _items[index];
            }
            set
            {
                try
                {
                    this[index] = (T)((object)value);
                }
                catch (InvalidCastException)
                {
                    throw new NotSupportedException();

                }
            }
        }

        bool IList.IsReadOnly
        {
            get
            {
                return _items.IsReadOnly;
            }
        }

        bool IList.IsFixedSize
        {
            get
            {
                IList list = _items as IList;
                if (list != null)
                    return list.IsFixedSize;
                return _items.IsReadOnly;
            }
        }

        int IList.Add(object value)
        {
            if (_items.IsReadOnly)
                throw new NotSupportedException();

            IfNullAndNullsAreIllegalThenThrow<T>(value);
            try
            {
                this.Add((T)((object)value));
            }
            catch (InvalidCastException)
            {
                throw new NotSupportedException();
            }
            return this.Count - 1;
        }

        bool IList.Contains(object value)
        {
            return Collection<T>.IsCompatibleObject(value) && this.Contains((T)((object)value));
        }

        int IList.IndexOf(object value)
        {
            if (Collection<T>.IsCompatibleObject(value))
            {
                return this.IndexOf((T)((object)value));
            }
            return -1;
        }

        void IList.Insert(int index, object value)
        {
            if (_items.IsReadOnly)
            {
                throw new NotSupportedException();
            }

            IfNullAndNullsAreIllegalThenThrow<T>(value);
            try
            {
                this.Insert(index, (T)((object)value));
            }
            catch (InvalidCastException)
            {
                throw new NotSupportedException();
            }
        }

        void IList.Remove(object value)
        {
            if (_items.IsReadOnly)
            {
                throw new NotSupportedException();
            }
            if (Collection<T>.IsCompatibleObject(value))
            {
                this.Remove((T)((object)value));
            }
        }

        #endregion

        #region IEnumerable

        public IEnumerator<T> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        #endregion

        #region Virtuals Methods

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        protected virtual void ClearItems()
        {
            _items.Clear();
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        protected virtual void InsertItem(int index, T item)
        {
            _items.Insert(index, item);
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        protected virtual void RemoveItem(int index)
        {
            _items.RemoveAt(index);
        }

        protected virtual void SetItem(int index, T item)
        {
            _items[index] = item;
        }

        #endregion

        #region Public Methods

        public void Add(T item)
        {
            if (_items.IsReadOnly)
                throw new NotSupportedException();
            int count = _items.Count;
            this.InsertItem(count, item);
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        public void Clear()
        {
            if (_items.IsReadOnly)
                throw new NotSupportedException();
            this.ClearItems();
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        public void CopyTo(T[] array, int index)
        {
            _items.CopyTo(array, index);
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        public bool Contains(T item)
        {
            return _items.Contains(item);
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        public int IndexOf(T item)
        {
            return _items.IndexOf(item);
        }

        public void Insert(int index, T item)
        {
            if (_items.IsReadOnly)
                throw new NotSupportedException();

            if (index < 0 || index > _items.Count)
                throw new NotSupportedException();

            this.InsertItem(index, item);
        }

        public bool Remove(T item)
        {
            if (_items.IsReadOnly)
                throw new NotSupportedException();

            int num = _items.IndexOf(item);
            if (num < 0)
                return false;

            this.RemoveItem(num);
            return true;
        }

        public void RemoveAt(int index)
        {
            if (_items.IsReadOnly)
                throw new NotSupportedException();
            if (index < 0 || index >= _items.Count)
                throw new NotSupportedException();

            this.RemoveItem(index);
        }

        #endregion

        #region Private Static Methods
       
        private static void IfNullAndNullsAreIllegalThenThrow<TT>(object value)
        {
            if (value == null && default(TT) != null)
            {
                throw new NotSupportedException();
            }
        }

        private static bool IsCompatibleObject(object value)
        {
            return value is T || (value == null && default(T) == null);
        }

        #endregion
    }
}

using System;
using System.ComponentModel;
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

        protected internal bool IsLastInsertItemSucseed { get; set; }

        protected internal bool IsLastRemoveItemSucseed { get; set; }

        protected internal virtual void OnGetCount()
        { 
        }

        public int Count
        {
            [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
            get
            {
                OnGetCount();
                return _items.Count;
            }
        }

        public virtual T this[int index]
        {
            [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
            get
            {
                T item = OnGetThisIndexerItem(index);
                if (null != item)
                    return item;
                item = RaiseGetThisIndexerItem(index);
                if (null != item)
                    return item;

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
                T item = OnGetThisIndexerItem(index);
                if (null != item)
                    return item;
                item = RaiseGetThisIndexerItem(index);
                if (null != item)
                    return item;
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

        public virtual IEnumerator<T> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        #endregion

        #region Virtual Methods

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        protected virtual void ClearItems()
        {
            _items.Clear();
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        protected virtual void InsertItem(int index, T item)
        {
            IsLastInsertItemSucseed = false;

            bool cancel = false;
            OnBeforeAddInsert(item, index, ref cancel);
            if (cancel)
                return;
            RaiseBeforeAddInsert(item, index, ref cancel);
            if (cancel)
                return;

            _items.Insert(index, item);

            OnAfterAddInsert(item, index);
            RaiseAfterAddInsert(item, index);

            IsLastInsertItemSucseed = true;
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        protected virtual void InsertItemSilent(int index, T item)
        {
            IsLastInsertItemSucseed = false;
            _items.Insert(index, item);
            this.FireListChanged(ListChangedType.ItemAdded, index);
            IsLastInsertItemSucseed = true;
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        protected virtual void RemoveItem(int index)
        {
            IsLastRemoveItemSucseed = false;
            T item = this[index];

            bool cancel = false;
            OnBeforeRemove(item, index, ref cancel);
            if (cancel)
                return;
            RaiseBeforeRemove(item, index, ref cancel);
            if (cancel)
                return;

            _items.RemoveAt(index);

            OnAfterRemove(item, index);
            RaiseAfterRemove(item, index);

            IsLastRemoveItemSucseed = true;
        }

        [TargetedPatchingOptOut("Performance critical to inline across NGen image boundaries")]
        protected virtual void RemoveItemSilent(int index)
        {
            IsLastRemoveItemSucseed = false;
            T item = this[index];

            _items.RemoveAt(index);

            IsLastRemoveItemSucseed = true;
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

        public void AddSilent(T item)
        {
            if (_items.IsReadOnly)
                throw new NotSupportedException();
            int count = _items.Count;
            this.InsertItemSilent(count, item);
            this.FireListChanged(ListChangedType.ItemAdded, count);
        }

        protected internal virtual void FireListChanged(ListChangedType type, int index)
        { 

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
            return IsLastRemoveItemSucseed;
        }

        public bool RemoveSilent(T item)
        {
            if (_items.IsReadOnly)
                throw new NotSupportedException();

            int num = _items.IndexOf(item);
            if (num < 0)
                return false;

            this.RemoveItemSilent(num);
            this.FireListChanged(ListChangedType.ItemDeleted, num);

            return IsLastRemoveItemSucseed;
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

        #region New Events

        public delegate void BeforeAddInsertEventHandler(T item, int itemIndex, ref bool cancel);

        public event BeforeAddInsertEventHandler BeforeAddInsert;

        private void RaiseBeforeAddInsert(T item, int itemIndex, ref bool cancel)
        {
            if (null != BeforeAddInsert)
                BeforeAddInsert(item, itemIndex, ref cancel);
        }

        
        public delegate void AfterAddInsertEventHandler(T item, int itemIndex);

        public event AfterAddInsertEventHandler AfterAddInsert;

        private void RaiseAfterAddInsert(T item, int itemIndex)
        {
            if (null != AfterAddInsert)
                AfterAddInsert(item, itemIndex);
        }


        public delegate void BeforeRemoveEventHandler(T item, int itemIndex, ref bool cancel);

        public event BeforeRemoveEventHandler BeforeRemove;

        private void RaiseBeforeRemove(T item, int itemIndex, ref bool cancel)
        {
            if (null != BeforeRemove)
                BeforeRemove(item, itemIndex, ref cancel);
        }


        public delegate void AfterRemoveEventHandler(T item, int itemIndex);

        public event AfterRemoveEventHandler AfterRemove;

        private void RaiseAfterRemove(T item, int itemIndex) 
        {
            if (null != AfterRemove)
                AfterRemove(item, itemIndex);
        }


        public delegate T GetThisIndexerItemEventHandler(int itemIndex);

        public event GetThisIndexerItemEventHandler GetThisIndexerItem;

        private T RaiseGetThisIndexerItem(int index)
        {
            if (null != GetThisIndexerItem)
                return GetThisIndexerItem(index);
            else
                return default(T);            
        }

        #endregion

        #region New Additional Virtuals

        /// <summary>
        /// Called before in this indexer _get 
        /// </summary>
        /// <param name="index">index of target item</param>
        /// <returns>item instance or null. if null the this indexer logic proceed normaly</returns>
        public virtual T OnGetThisIndexerItem(int index)
        {
            return default(T);
        }

        /// <summary>
        /// Called before an item was removed
        /// </summary>
        /// <param name="item">the target item to delete</param>
        /// <param name="itemIndex">the index of the target item</param>
        /// <param name="cancel">cancel operation flag</param>
        protected internal virtual void OnBeforeRemove(T item, int itemIndex, ref bool cancel)
        {

        }

        /// <summary>
        /// Called after an item was removed
        /// </summary>
        /// <param name="item">the deleted item. WARNING: the item is not a list member anymore</param>
        protected internal virtual void OnAfterRemove(T item, int index)
        {

        }

        /// <summary>
        /// Called before a new item was added or inserted
        /// </summary>
        /// <param name="item">the target item to add. WARNING: the item is not a list member currently</param>
        /// <param name="itemIndex">the new index for the target item</param>
        /// <param name="cancel">cancel operation flag</param>
        protected internal virtual void OnBeforeAddInsert(T item, int itemIndex, ref bool cancel)
        {

        }    

        /// <summary>
        /// Called after an item was added
        /// </summary>
        /// <param name="item">the added item. </param>
        /// <param name="itemIndex">the new index for the added item</param>
        protected internal virtual void OnAfterAddInsert(T item, int itemIndex)
        {

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

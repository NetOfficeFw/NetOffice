using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{   
    /// <summary>
    /// Handles local changes for an AccessContextList instance
    /// </summary>
    internal class AccessContextListItemManager
    {
        #region Embedded Types

        internal enum ItemEntryState
        { 
            New = 0,
            Deleted = 1,
            Changed = 2
        }

        internal class ItemEntry
        {
            internal ItemEntry(AccessContextItem item, ItemEntryState state)
            {
                Item = item;
                State = state;
            }

            public ItemEntryState State { get; private set; }
            public AccessContextItem Item { get; private set; }
        }

        internal class ItemIndexDecrement
        {
            internal ItemIndexDecrement(ItemEntry item)
            {
                Item = item;
            }

            public ItemEntry Item { get; private set; }
            public int DecrementValue { get; internal set; }
        }

        #endregion

        #region Fields

        /// <summary>
        /// decrement index memory. sometimes its necessary to decrement the index information for items in the manager instance.
        /// its impossible to change Dictionary KeyValuePair values. this is the reason we use a memory.
        /// The AccessContextList=>UpdateFromOtherInstance method stores decrement info for manager instance items.
        /// The AccessContextList=>CancelLocalChanges method is aware of the decrement memory and use them.
        /// </summary>
        private List<ItemIndexDecrement> _decrementMemory = new List<ItemIndexDecrement>();

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">parent list</param>
        public AccessContextListItemManager(AccessContextList parent)
        {
            Parent = parent;
            Items = new Dictionary<ItemEntry, int>();
        }

        #endregion

        #region Properties
        
        /// <summary>
        /// Parent list
        /// </summary>
        public AccessContextList Parent { get; private set; }

        /// <summary>
        /// Count of all local deleted items
        /// </summary>
        public int DeletedCount
        {
            get
            {
                int count = 0;
                foreach (var item in Items)
                    if (item.Key.State == ItemEntryState.Deleted)
                        count++;
                return count;
            }
        }

        /// <summary>
        /// Count of all local new created items
        /// </summary>
        public int NewCount
        {
            get
            {
                int count = 0;
                foreach (var item in Items)
                    if (item.Key.State == ItemEntryState.New)
                        count++;
                return count;
            }
        }
        
        /// <summary>
        /// Count of all local changed items, includes not deleted and new items
        /// </summary>
        public int ChangedCount
        {
            get
            {
                int count = 0;
                foreach (var item in Items)
                    if (item.Key.State == ItemEntryState.Changed)
                        count++;
                return count;
            }
        }

        /// <summary>
        /// Returns info the item manager contains local changes
        /// </summary>
        public bool ContainsLocalChanges
        {
            get 
            {
                return Items.Count > 0;
            }
        }
        
        /// <summary>
        /// All new/deleted/changed items ItemEntry=>Item, int=> index of the item
        /// </summary>
        internal Dictionary<ItemEntry, int> Items { get; private set; }

        #endregion

        #region Methods

        /// <summary>
        /// Clears the manager instance
        /// </summary>
        internal void Clear()
        {
            Items.Clear();
            _decrementMemory.Clear();
            Parent.RaiseNotifyPropertyChanged("ContainsLocalChanges");
        }

        /// <summary>
        /// Returns info, an item is local new created
        /// </summary>
        /// <param name="item">target item</param>
        /// <returns>true if local new created otherwise false</returns>
        public bool IsNewItem(AccessContextItem item)
        {
            foreach (var listItem in Items)
                if (listItem.Key.Item == item && listItem.Key.State == ItemEntryState.New)
                    return true;
            return false;
        }

        /// <summary>
        /// Returns info, an item is local deleted
        /// </summary>
        /// <param name="item">target item</param>
        /// <returns>true if local deleted otherwise false</returns>
        public bool IsDeletedItem(AccessContextItem item)
        {
            foreach (var listItem in Items)
                if (listItem.Key.Item == item && listItem.Key.State == ItemEntryState.Deleted)
                    return true;
            return false;
        }

        /// <summary>
        /// Returns info, an item is local changed (not new or deleted)
        /// </summary>
        /// <param name="item">target item</param>
        /// <returns>true if local changed otherwise false</returns>
        public bool IsChangedItem(AccessContextItem item)
        {
            foreach (var listItem in Items)
                if (listItem.Key.Item == item && listItem.Key.State == ItemEntryState.Changed)
                    return true;
            return false;
        }

        /// <summary>
        /// Add item to the manager instance and mark as local changed
        /// </summary>
        /// <param name="item">target item</param>
        /// <param name="index">current index of target item</param>
        public void AddChangedItem(AccessContextItem item, int index)
        {
            if (item.ItemState != AccessContextItemState.ItemIsDeleted && !IsChangedItem(item) && !(IsNewItem(item)))
            {
                item.ItemState = AccessContextItemState.ItemIsLocalChanged;
                Items.Add(new ItemEntry(item, ItemEntryState.Changed), index);
                Parent.RaiseNotifyPropertyChanged("ContainsLocalChanges");
            }
        }
       
        /// <summary>
        /// Add item to the manager instance and mark as local new created
        /// </summary>
        /// <param name="item">target item</param>
        /// <param name="index">current index of target item</param>
        public void AddNewItem(AccessContextItem item, int index)
        {
            if (!IsNewItem(item))
            {
                Items.Add(new ItemEntry(item, ItemEntryState.New), index);
                Parent.RaiseNotifyPropertyChanged("ContainsLocalChanges");
            }
        }

        /// <summary>
        /// Add item to the manager instance and mark as local deleted
        /// </summary>
        /// <param name="item"></param>
        public void AddDeletedItem(AccessContextItem item)
        {
            if (!IsDeletedItem(item))
            {
                int currentIndex = item.Parent.IndexOf(item);
                Items.Add(new ItemEntry(item, ItemEntryState.Deleted), currentIndex);
                item.ItemState = AccessContextItemState.ItemIsLocalDeleted;
                Parent.RaiseNotifyPropertyChanged("ContainsLocalChanges");
            }
        }

        /// <summary>
        /// Remove an item from the manager instance that is marked as local new created
        /// </summary>
        /// <param name="item">target item</param>
        public void RemoveNewItem(AccessContextItem item)
        {
            if (IsNewItem(item))
            {
                ItemEntry entry = null;
                foreach (var listItem in Items)
                {
                    if (listItem.Key.Item == item && listItem.Key.State == ItemEntryState.New)
                    {
                        entry = listItem.Key;
                        break;
                    }
                }
                if (null != entry)
                {
                    Console.WriteLine("RemoveNewItem " + item.ToString());
                    Items.Remove(entry);
                    item.ItemState = AccessContextItemState.ItemIsDeleted;
                    Parent.RaiseNotifyPropertyChanged("ContainsLocalChanges");
                }
            }
        }
       
        /// <summary>
        /// Removes an new/deleted/changed item
        /// </summary>
        /// <param name="item"></param>
        internal void RemoveItem(AccessContextItem item)
        {
            ItemEntry entry = null;
            foreach (var listItem in Items)
            {
                if (listItem.Key.Item == item)
                {
                    entry = listItem.Key;
                    break;
                }
            }

            if (null != entry)
            {
                Items.Remove(entry);
                Parent.RaiseNotifyPropertyChanged("ContainsLocalChanges");
            }
        }

        /// <summary>
        /// Get an item from associated data source
        /// </summary>
        /// <param name="item">data source item</param>
        /// <returns>AccessContextItem instance or null</returns>
        internal AccessContextItem GetItemFromDataSource(RootItem item)
        {
            foreach (var listItem in Items)
            {
                if (listItem.Key.Item.DataSource == item)
                    return listItem.Key.Item;
            }
            return null;
        }

        /// <summary>
        /// Get Mananger instance Items-Dictionary as Array in reverse order
        /// </summary>
        /// <returns>Items array. ItemEntry=>Item, int=>index </returns>
        internal KeyValuePair<ItemEntry, int>[] GetItemsInReverseOrder()
        {
            List<ItemEntry> list = new List<ItemEntry>();
            foreach (var item in Items)
                list.Insert(0, item.Key);

            List<KeyValuePair<ItemEntry, int>> resultList = new List<KeyValuePair<ItemEntry, int>>();

            foreach (var item in list)
            {
                int index = Items[item];
                resultList.Add(new KeyValuePair<ItemEntry, int>(item, index));
            }

            return resultList.ToArray();
        }

        /// <summary>
        /// Decrement all indexes for items in the manager instances
        /// </summary>
        /// <param name="equalOrAboveStart">items with an index that is equal or above to this argument</param>
        internal void DecrementItemIndex(int equalOrAboveStart)
        {
            foreach (var item in Items)
            {
                if (item.Key.State == ItemEntryState.Deleted && item.Value >= equalOrAboveStart)
                {
                    ItemIndexDecrement itemDecrement = GetIndexDecrement(item.Key, true);
                    itemDecrement.DecrementValue++;
                }
            }
        }

        /// <summary>
        /// Get an index decrement info instance
        /// </summary>
        /// <param name="entry">target item</param>
        /// <param name="autoCreateNew">auto create new if not exists</param>
        /// <returns>ItemIndexDecrement instance or null(if not autoCreateNew)</returns>
        internal ItemIndexDecrement GetIndexDecrement(ItemEntry entry, bool autoCreateNew)
        {
            foreach (var item in _decrementMemory)
            {
                if (item.Item == entry)
                    return item;
            }

            if (autoCreateNew)
            {
                ItemIndexDecrement newDecrement = new ItemIndexDecrement(entry);
                _decrementMemory.Add(newDecrement);
                return newDecrement;
            }
            else
                return null;
        }

        #endregion
    }
}

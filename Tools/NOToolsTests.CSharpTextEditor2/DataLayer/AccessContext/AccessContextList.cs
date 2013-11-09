using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CM = NOTools.ComponentModel;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Proxy table for a RootList instance
    /// </summary>
    public class AccessContextList : CM.BindingList<AccessContextItem>, ITypedList, INotifyPropertyChanged
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">associated context</param>
        /// <param name="dataSource">origin table</param>
        public AccessContextList(AccessContext parent, RootList dataSource)
        {
            Parent = parent;
            DataSource = dataSource;
            ItemManager = new AccessContextListItemManager(this);
            ResetLocalData();
        }
        
        #endregion

        #region Properties

        /// <summary>
        /// Handles local new created/deleted items
        /// </summary>
        private AccessContextListItemManager ItemManager { get; set; }

        /// <summary>
        /// Associated context
        /// </summary>
        private AccessContext Parent { get; set; }

        /// <summary>
        /// Origin table
        /// </summary>
        internal RootList DataSource { get; set; }

        /// <summary>
        /// Name of the Table/DataSource
        /// </summary>
        public string Name
        {
            get 
            {
                return DataSource.Name;
            }
        }

        /// <summary>
        /// Returns info the proxy table instance contains local changes
        /// </summary>
        public bool ContainsLocalChanges
        {
            get 
            {
                return ItemManager.ContainsLocalChanges;
            }
        }

        #endregion

        #region INotifyPropertyChanged

        /// <summary>
        /// Occurs when the ContainsLocalChanges property has changed
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        internal void RaiseNotifyPropertyChanged(string propertyName)
        {
            if (null != PropertyChanged)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
                Parent.RaiseNotifyPropertyChanged(propertyName);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Mark item as local changed. The method was called from item himself while after a property value has changed
        /// </summary>
        /// <param name="item"></param>
        /// <param name="propertyName"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        internal void MarkItemAsLocalChanged(AccessContextItem item, string propertyName, object oldValue, object newValue)
        {
            ItemManager.AddChangedItem(item, IndexOf(item));
            Parent.Item_IsLocalChanged(item, propertyName, oldValue, newValue);
        }

        /// <summary>
        /// Reload data from root tables
        /// </summary>
        public void ResetLocalData()
        {
            ItemManager.Clear();
            this.Clear();
            foreach (RootItem item in DataSource)
                AddSilent(new AccessContextItem(this, item, AccessContextItemState.ItemIsNormal));
            FireListChanged(ListChangedType.Reset, -1);
        }
        
        /// <summary>
        /// Rollback local changes without reload data from root table
        /// </summary>
        internal void CancelLocalChanges()
        {
            foreach (var item in ItemManager.GetItemsInReverseOrder())
            {
                switch (item.Key.State)
                {
                    case AccessContextListItemManager.ItemEntryState.New:

                        item.Key.Item.ItemState = AccessContextItemState.ItemIsDeleted;
                        RemoveSilent(item.Key.Item);

                        break;
                    case AccessContextListItemManager.ItemEntryState.Deleted:

                        item.Key.Item.ItemState = AccessContextItemState.ItemIsNormal;
                        int decrement = 0;
                        var v = ItemManager.GetIndexDecrement(item.Key, false);
                        if (null != v)
                            decrement = v.DecrementValue;
                        InsertItemSilent(item.Value - decrement, item.Key.Item);
                        item.Key.Item.LocalChangedProperties.Clear();

                        break;
                    case AccessContextListItemManager.ItemEntryState.Changed:

                        item.Key.Item.ItemState = AccessContextItemState.ItemIsNormal;
                        item.Key.Item.LocalChangedProperties.Clear();
                        FireListChanged(ListChangedType.ItemChanged, item.Value);

                        break;
                    default:
                        throw new IndexOutOfRangeException();
                }
            }
            ItemManager.Clear();
        }
        
        /// <summary>
        /// Commit local changes to associated root table
        /// </summary>
        internal void ApplyLocalChanges()
        {
            foreach (var item in ItemManager.Items)
            {
                switch (item.Key.State)
                {
                    case AccessContextListItemManager.ItemEntryState.New:

                        item.Key.Item.ItemState = AccessContextItemState.ItemIsNormal;
                        RootItem newItem = DataSource.AddNew();
                        foreach (var property in item.Key.Item.LocalChangedProperties)
                            newItem.SetValue(property.Key, property.Value);
                        item.Key.Item.DataSource = newItem;

                        // datenbank kommando
                        // dynmamischer code
                        break;
                    case AccessContextListItemManager.ItemEntryState.Deleted:

                        item.Key.Item.ItemState = AccessContextItemState.ItemIsDeleted;
                        DataSource.Remove(item.Key.Item.DataSource);

                        // datenbank kommando
                        // dynmamischer code
                        break;
                    case AccessContextListItemManager.ItemEntryState.Changed:

                        item.Key.Item.ItemState = AccessContextItemState.ItemIsNormal;
                        foreach (var property in item.Key.Item.LocalChangedProperties)
                            item.Key.Item.DataSource.SetValue(property.Key, property.Value);

                        // datenbank kommando
                        // dynmamischer code
                        break;
                    default:
                        throw new IndexOutOfRangeException();
                }
            }
            Parent.Parent.UpdateNotifyOtherListInstances(this);
            ItemManager.Clear();
        }

        /// <summary>
        /// Update the instance data from another AccessContextList instance.
        /// The method was called after commit changes from another access context to synchronize view data.
        /// </summary>
        /// <param name="instance">other access context list from different access context</param>
        internal void UpdateFromOtherInstance(AccessContextList instance)
        {
            if (this == instance)
                throw new InvalidOperationException();

            foreach (var item in instance.ItemManager.Items)
            {
                switch (item.Key.State)
                {
                    case AccessContextListItemManager.ItemEntryState.New:

                        int newIndex = this.Count - ItemManager.NewCount;
                        AccessContextItem newContextItem = new AccessContextItem(this, item.Key.Item.DataSource, AccessContextItemState.ItemIsNormal);
                        InsertItemSilent(newIndex, newContextItem);
                       
                        break;
                    case AccessContextListItemManager.ItemEntryState.Deleted:

                        AccessContextItem deleteItem = GetItemFromDataSource(item.Key.Item.DataSource);
                        if (null == deleteItem)
                        {
                            deleteItem = ItemManager.GetItemFromDataSource(item.Key.Item.DataSource);
                            if (null != deleteItem)
                            { 
                                deleteItem.ItemState = AccessContextItemState.ItemIsDeleted;
                                ItemManager.RemoveItem(deleteItem);
                            }
                            continue;
                        }
                        deleteItem.ItemState = AccessContextItemState.ItemIsDeleted;
                        int deleteIndex = IndexOf(deleteItem);
                        RemoveSilent(deleteItem);
                        ItemManager.DecrementItemIndex(deleteIndex);

                        break;
                    case AccessContextListItemManager.ItemEntryState.Changed:

                        // do nothing

                        break;
                    default:
                        throw new IndexOutOfRangeException();
                }
            }
        }
    
        /// <summary>
        /// Try to get an AccessContextItem instance from associated data source
        /// </summary>
        /// <param name="item">data source item</param>
        /// <returns>AccessContextItem instance or null</returns>
        private AccessContextItem GetItemFromDataSource(RootItem item)
        {
            foreach (var listItem in this)
            {
                if (listItem.DataSource == item)
                    return listItem;
            }
            return null;
        }

        #endregion

        #region ITypedList

        /// <summary>
        /// Returns the property descriptors for the list instance
        /// </summary>
        /// <param name="listAccessors">target properties. unused argument!</param>
        /// <returns>property descriptors</returns>
        public PropertyDescriptorCollection GetItemProperties(PropertyDescriptor[] listAccessors)
        {
            List<PropertyDescriptor> list = new List<PropertyDescriptor>();

            foreach (PropertyDescriptor item in DataSource.GetItemProperties(listAccessors))
                list.Add(new AccessContextPropertyDescriptor(item));

            return new PropertyDescriptorCollection(list.ToArray());
        }

        /// <summary>
        /// Returns the name of the List instance
        /// </summary>
        /// <param name="listAccessors">target properties. unused argument!</param>
        /// <returns>System.String</returns>
        public string GetListName(PropertyDescriptor[] listAccessors)
        {
            return DataSource.GetListName(listAccessors);
        }

        #endregion
        
        #region Overrides

        protected override void OnBeforeRemove(AccessContextItem item, int itemIndex, ref bool cancel)
        {
            if (ItemManager.IsNewItem(item))
                ItemManager.RemoveNewItem(item);
            else
                ItemManager.AddDeletedItem(item);
        }

        protected override void OnBeforeAddInsert(AccessContextItem item, int itemIndex, ref bool cancel)
        {
            if (!ItemManager.IsNewItem(item))
                ItemManager.AddNewItem(item, itemIndex);
        }

        protected override void ResolveArgumentsOnCreateNew(ref object[] args)
        {
            args = new object[] { this, AccessContextItemState.ItemIsNew};
        }

        #endregion
    }
}

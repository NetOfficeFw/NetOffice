using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;
using System.Runtime;
using System.Collections;
using NetOffice.Tools;

namespace NetOffice.OfficeApi.Tools.Utils
{  
    /// <summary>
    /// Represents a collection of tray menu items
    /// </summary>
    public class TrayMenuItems : IEnumerable<TrayMenuItem>
    {
        #region Fields

        private List<TrayMenuItem> _items = new List<TrayMenuItem>();

        private TrayMenu _owner;

        private TrayMenuItem _parent;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">collection owner</param>
        internal TrayMenuItems(TrayMenu owner)
        {
            if (null == owner)
                throw new ArgumentNullException();
            _owner = owner;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">collection owner</param>
        /// <param name="parent">parent item instance</param>
        internal TrayMenuItems(TrayMenu owner, TrayMenuItem parent)
        {
            if (null == owner)
                throw new ArgumentNullException();
            _owner = owner;
            _parent = parent;
        }

        #endregion

        #region Events

        /// <summary>
        /// Occurs after item has been added
        /// </summary>
        public event TrayMenuItemsChangedHandler ItemAdded;

        /// <summary>
        /// Occurs after item has been removed
        /// </summary>
        public event TrayMenuItemsChangedHandler ItemRemoved;

        /// <summary>
        /// Occurs after instance has been cleared
        /// </summary>
        public event EventHandler ItemsClear;

        #endregion

        #region Properties
        
        /// <summary>
        /// Returns the item count
        /// </summary>
        public virtual int Count
        {
            get
            {
                return _items.Count;
            }
        }

        /// <summary>
        /// Returns item from specified index
        /// </summary>
        /// <param name="index">target</param>
        /// <returns>item from index</returns>
        public virtual TrayMenuItem this[int index]
        {
            get
            {
                return _items[index];
            }
        }

        /// <summary>
        /// Parent Item Instance
        /// </summary>
        protected internal TrayMenuItem Parent
        {
            get
            {
                return _parent;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Add items to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <param name="control">custom control</param>
        /// <returns>new created items</returns>
        public virtual T Add<T>(string text, object control) where T : TrayMenuItem
        {
            TrayMenuItemType type = GetItemType<T>();
            return Add(text, true, null, type, control) as T;
        }


        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <returns>new created items</returns>
        public virtual T Add<T>() where T : TrayMenuItem
        {
            TrayMenuItemType type = GetItemType<T>();
            return Add(String.Empty, true, null, type) as T;
        }

        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <returns>new created items</returns>
        public virtual T Add<T>(string text) where T : TrayMenuItem
        {
            TrayMenuItemType type = GetItemType<T>();
            return Add(text, true, null, type) as T;
        }

        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <param name="visible">item visibility</param>
        /// <returns>new created item</returns>
        public virtual T Add<T>(string text, bool visible) where T : TrayMenuItem
        {
            TrayMenuItemType type = GetItemType<T>();
            return Add(text, visible, null, type) as T;
        }

        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <param name="visible">item visibility</param>
        /// <param name="image">item image</param>
        /// <returns>new created item</returns>
        public virtual T Add<T>(string text, bool visible, Image image) where T : TrayMenuItem
        {
            TrayMenuItemType type = GetItemType<T>();
            return Add(text, visible, image, type) as T;
        }

        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <param name="visible">item visibility</param>
        /// <param name="imageResourceName">image resource address</param>
        /// <returns>new created item</returns>
        public virtual T Add<T>(string text, bool visible, string imageResourceName) where T : TrayMenuItem
        {
            TrayMenuItemType type = GetItemType<T>();
            return Add(text, visible, Image.FromStream(ReadRessource(imageResourceName)), type) as T;
        }
        
        /// <summary>
        /// Add items to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <returns>new created items</returns>
        public virtual IEnumerable<T> Add<T>(params string[] text) where T : TrayMenuItem
        {
            // to validate T before
            GetItemType<T>();

            T[] result = new T[text.Length];

            for (int i = 0; i < text.Length; i++)
            {
                string itemText = text[i];
                T item = Add<T>(itemText);          
                result[i] = item;             
            }

            return result;
        }

        /// <summary>
        /// Add items to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <returns>new created items</returns>
        public virtual IEnumerable<TrayMenuItem> Add(params string[] text)
        {
            TrayMenuItem[] result = new TrayMenuItem[text.Length];

            for (int i = 0; i < text.Length; i++)
            {
                string itemText = text[i];
                TrayMenuItem item = new TrayMenuItem(_owner, itemText);
                _items.Add(item);
                result[i] = item;
                RaiseItemAdded(item);
            }            

            return result;
        }
        
        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <returns>new created item</returns>
        public virtual TrayMenuItem Add(string text)
        {           
            TrayMenuItem item = new TrayMenuItem(_owner, text);
            _items.Add(item);
            RaiseItemAdded(item);
            return item;
        }

        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <param name="visible">item visibility</param>
        /// <returns>new created item</returns>
        public virtual TrayMenuItem Add(string text, bool visible)
        {          
            TrayMenuItem item = new TrayMenuItem(_owner, text, visible);
            _items.Add(item);
            RaiseItemAdded(item);
            return item;
        }

        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <param name="visible">item visibility</param>
        /// <param name="image">item image</param>
        /// <returns>new created item</returns>
        public virtual TrayMenuItem Add(string text, bool visible, Image image)
        {
            TrayMenuItem item = new TrayMenuItem(_owner, text, visible);
            item.Image = image;
            _items.Add(item);
            RaiseItemAdded(item);
            return item;
        }

        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <param name="itemType">new item type</param>
        /// <returns>new created item</returns>
        public virtual TrayMenuItem Add(string text, TrayMenuItemType itemType)
        {
            return Add(text, true, null, itemType);
        }

        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <param name="visible">item visibility</param>
        /// <param name="itemType">new item type</param>
        /// <returns>new created item</returns>
        public virtual TrayMenuItem Add(string text, bool visible, TrayMenuItemType itemType)
        {
            return Add(text, visible, null, itemType);
        }
       
        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <param name="visible">item visibility</param>
        /// <param name="image">item image</param>
        /// <param name="itemType">new item type</param>
        /// <returns>new created item</returns>
        public virtual TrayMenuItem Add(string text, bool visible, Image image, TrayMenuItemType itemType)
        {
            return Add(text, visible, image, itemType, null);
        }

        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="text">shown item caption</param>
        /// <param name="visible">item visibility</param>
        /// <param name="image">item image</param>
        /// <param name="itemType">new item type</param>
        ///  <param name="control">custom control</param>
        /// <returns>new created item</returns>
        public virtual TrayMenuItem Add(string text, bool visible, Image image, TrayMenuItemType itemType, object control)
        {          
            TrayMenuItem item = null;

            switch (itemType)
            {
                case TrayMenuItemType.Item:
                    item = new TrayMenuItem(_owner, text, visible);
                    break;
                case TrayMenuItemType.Label:
                    item = new TrayMenuLabelItem(_owner, text, visible);
                    break;
                case TrayMenuItemType.LinkLabel:
                    item = new TrayMenuLinkLabelItem(_owner, text, visible);
                    break;
                case TrayMenuItemType.Button:
                    item = new TrayMenuButtonItem(_owner, text, visible);
                    break;
                case TrayMenuItemType.TextBox:
                    item = new TrayMenuTextboxItem(_owner, text, visible);
                    break;
                case TrayMenuItemType.CheckBox:
                    item = new TrayMenuCheckboxItem(_owner, text, visible);
                    break;
                case TrayMenuItemType.Progress:
                    item = new TrayMenuProgressItem(_owner, text, visible);
                    break;
                case TrayMenuItemType.DropDownList:
                    item = new TrayMenuDropDownListItem(_owner, text, visible);
                    break;
                case TrayMenuItemType.Separator:
                    item = new TrayMenuSeparatorItem(_owner, text, visible);
                    break;
                case TrayMenuItemType.Custom:
                    if (!(control is System.Windows.Forms.Control))
                        throw new ArgumentOutOfRangeException("control");
                    item = new TrayMenuCustomItem(_owner, text, visible, control);
                    break;
                case TrayMenuItemType.Monitor:
                    item = new TrayMenuMonitorItem(_owner, text, visible);
                    break;
                case TrayMenuItemType.AutoClose:
                    item = new TrayMenuAutoCloseItem(_owner, text, visible);
                    break;
                case TrayMenuItemType.Close:
                    item = new TrayMenuCloseItem(_owner, text, visible);
                    break;
                default:
                    throw new ArgumentOutOfRangeException("itemType");
            }

            item.Image = image;
            _items.Add(item);
            RaiseItemAdded(item);
            return item;
        }
        
        /// <summary>
        /// Remove an item 
        /// </summary>
        /// <param name="item">target item to remove</param>
        /// <returns>true if its removed, otherwise false</returns>
        public virtual bool Remove(TrayMenuItem item)
        {
            int itemIndex = IndexOf(item);
            bool result = _items.Remove(item);
            if (result)
                RaiseItemRemoved(item, itemIndex);
            return result;
        }

        /// <summary>
        /// Remove an item from specified index
        /// </summary>
        /// <param name="index">target index to remove</param>
        public virtual void RemoveAt(int index)
        {
            TrayMenuItem item = _items[index];
            _items.RemoveAt(index);
            RaiseItemRemoved(item, index);
        }

        /// <summary>
        /// Remove all items
        /// </summary>
        public virtual void Clear()
        {
            if (_items.Count > 0)
            {
                _items.Clear();
                RaiseItemsClear();
            }
        }

        /// <summary>
        /// Returns item index
        /// </summary>
        /// <param name="item">target item</param>
        /// <returns>target index</returns>
        public virtual int IndexOf(TrayMenuItem item)
        {
            if (null == item)
                throw new ArgumentNullException();
            return _items.IndexOf(item);
        }

        /// <summary>
        /// Raise the ItemAdded event
        /// </summary>
        /// <param name="item">new created item</param>
        private void RaiseItemAdded(TrayMenuItem item)
        {
            _owner.OnItemAdded(_parent, item);
            if (null != ItemAdded)
                ItemAdded(this, new TrayMenuItemsEventArgs(item));
        }

        /// <summary>
        /// Raise the ItemRemoved event
        /// </summary>
        /// <param name="item">removed item</param>
        /// <param name="itemIndex">item index</param>
        private void RaiseItemRemoved(TrayMenuItem item, int itemIndex)
        {
            _owner.OnItemRemoved(_parent, item, itemIndex);
            if (null != ItemRemoved)
                ItemRemoved(this, new TrayMenuItemsEventArgs(item));
        }

        /// <summary>
        /// Raise the ItemsClear event
        /// </summary>
        private void RaiseItemsClear()
        {
            _owner.OnItemsClear();
            if (null != ItemsClear)
                ItemsClear(this, EventArgs.Empty);
        }

        /// <summary>
        /// Returns information which item type represents T
        /// </summary>
        /// <typeparam name="T">item class or derived</typeparam>
        /// <returns>runtime supported item type</returns>
        private TrayMenuItemType GetItemType<T>() where T : TrayMenuItem
        {
            Type type = typeof(T);
            object[] attribs = type.GetCustomAttributes(typeof(ItemTypeAttribute), true);
            if (null == attribs || attribs.Length == 0)
                throw new ArgumentException();
            return (attribs[0] as ItemTypeAttribute).Type;
        }

        /// <summary>
        /// Read resource stream
        /// </summary>
        /// <param name="address">resource address</param>
        /// <returns>stream instance</returns>
        private System.IO.Stream ReadRessource(string address)
        {
            if (null == _owner || null == _owner.Owner)
                throw new System.IO.IOException("Unable to access resource stream.");

            System.Reflection.Assembly assembly = _owner.Owner.Type.Assembly;
            System.IO.Stream stream = assembly.GetManifestResourceStream(address);
            if (null == stream)
            {
                string space = _owner.Owner.GetType().Namespace;
                stream = assembly.GetManifestResourceStream(space + "." + address);
                return stream;
            }
            else
                throw new System.IO.IOException("Unable to find resource stream.");
        }

        #endregion

        #region IEnumerable<TrayMenuItem>

        /// <summary>
        /// Instance enumerator
        /// </summary>
        /// <returns>enumerator</returns>
        public virtual IEnumerator<TrayMenuItem> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        /// <summary>
        /// Instance enumerator
        /// </summary>
        /// <returns>enumerator</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        #endregion
    }
}

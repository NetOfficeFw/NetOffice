using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Text;
using System.Drawing;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents a tray menu drop down box item
    /// </summary>
    [ItemType(TrayMenuItemType.DropDownList)]
    public class TrayMenuDropDownListItem : TrayMenuItem
    {
        #region Nested

        /// <summary>
        /// DropDownList Display Style
        /// </summary>
        public enum DropDownListStyle
        {            
            /// <summary>
            /// Table View With Free Text
            /// </summary>
            Simple = 0,

            /// <summary>
            /// Drop Down View With Free Text
            /// </summary>
            DropDown = 1,

            /// <summary>
            /// Drop Down View With Fixed Selection
            /// </summary>
            DropDownList = 2
        }

        /// <summary>
        /// Represents the items in a TrayMenuDropDownListItem instance
        /// </summary>
        public class ObjectCollection : IEnumerable<object>
        {
            private List<object> _items = new List<object>();

            internal ObjectCollection(TrayMenuDropDownListItem parent)
            {
                if (null == parent)
                    throw new ArgumentNullException();
                Parent = parent;
            }

            private TrayMenuDropDownListItem Parent { get; set; }

            /// <summary>
            /// Count of items
            /// </summary>
            public int Count
            {
                get
                {
                    return _items.Count;
                }
            }

            /// <summary>
            /// Get or set item at the specified index
            /// </summary>
            /// <param name="index"></param>
            /// <returns></returns>
            public object this[int index]
            {
                get
                {
                    return _items[index];
                }
                set
                {
                    _items[index] = value;
                }
            }

            /// <summary>
            /// Add a new item to the collection
            /// </summary>
            /// <param name="item">new item</param>
            public void Add(object item)
            {
                _items.Add(item);
                Parent.Owner.OnDropDownItem_ListItemAdded(Parent, item);
            }

            /// <summary>
            /// Add new items to the collection
            /// </summary>
            /// <param name="items">new items</param>
            public void Add(params object[] items)
            {
                if (null == items)
                    return;
                foreach (var item in items)
                {
                    _items.Add(item);
                    Parent.Owner.OnDropDownItem_ListItemAdded(Parent, item);
                }               
            }

            /// <summary>
            /// Remove an item from the collection
            /// </summary>
            /// <param name="item">item to remove</param>
            /// <returns>true if removed, otherwise false</returns>
            public bool Remove(object item)
            {
                int itemIndex = _items.IndexOf(item);
                if (itemIndex > -1)
                {
                    _items.Remove(item);
                    Parent.Owner.OnDropDownItem_ListItemRemoved(Parent, item, itemIndex);
                    return true;
                }
                else
                    return false;
            }

            /// <summary>
            /// Removes all items
            /// </summary>
            public void Clear()
            {
                if (_items.Count > 0)
                {
                    _items.Clear();
                    Parent.Owner.OnDropDownItem_ListItemsCleared(Parent);
                }
            }

            /// <summary>
            /// Returns an enumerable item sequence
            /// </summary>
            /// <returns>enumerator</returns>
            public IEnumerator<object> GetEnumerator()
            {
                return _items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return _items.GetEnumerator();
            }
        }

        #endregion

        #region Fields

        private DropDownListStyle _dropdownStyle = DropDownListStyle.DropDown;
        private int _dropDownHeight;
        private int _maxLength;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        internal TrayMenuDropDownListItem(TrayMenu owner, string text) : base(owner, text)
        {
            ItemType = TrayMenuItemType.DropDownList;
            DataSource = new ObjectCollection(this);
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        /// <param name="visible">item visibility</param>
        internal TrayMenuDropDownListItem(TrayMenu owner, string text, bool visible) : base(owner, text, visible)
        {
            ItemType = TrayMenuItemType.DropDownList;
            DataSource = new ObjectCollection(this);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Shown items
        /// </summary>
        public ObjectCollection DataSource { get; private set; }

        /// <summary>
        /// DropDown List Style
        /// </summary>
        public DropDownListStyle DropDownStyle
        {
            get
            {
                return _dropdownStyle;
            }
            set
            {
                if (value != _dropdownStyle)
                {
                    _dropdownStyle = value;
                    Owner.OnDropDownItemStyleChanged(this);
                }
            }
        }

        /// <summary>
        ///Drop Down Height
        /// </summary>
        public int DropDownHeight
        {
            get
            {
                return _dropDownHeight;
            }
            set
            {
                _dropDownHeight = value;
                _dropDownHeight = Owner.OnDropDownItemDropDownHeightChanged(this);
            }
        }

        /// <summary>
        /// Max Text Length
        /// </summary>
        public int MaxLength
        {
            get
            {
                return _maxLength;
            }
            set
            {
                _maxLength = value;
                _maxLength = Owner.OnDropDownItemMaxLengthChanged(this);
            }
        }

        #endregion

        #region Methods

        internal void SetupDropDownElements(int dropDownHeight)
        {
            _dropDownHeight = dropDownHeight;
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Shown Image
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public override Image Image
        {
            get
            {
                return base.Image;
            }
            set
            {
                base.Image = value;
            }
        }
        /// <summary>
        /// Optional child items which is not supported in this item type
        /// </summary>
        [System.ComponentModel.Browsable(false), System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)]
        public override TrayMenuItems Items
        {
            get
            {
                return base.Items;
            }
        }
        /// <summary>
        /// Creates a new items collection
        /// </summary>
        /// <returns>collection instance</returns>
        protected internal override TrayMenuItems OnCreateMenuItems()
        {
            return new TrayMenuStubItems(Owner, this);
        }

        #endregion
    }
}

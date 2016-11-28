using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents a tray menu checkbox item
    /// </summary>
    [ItemType(TrayMenuItemType.CheckBox)]
    public class TrayMenuCheckboxItem : TrayMenuItem
    {
        #region Fields

        private bool _checked;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        internal TrayMenuCheckboxItem(TrayMenu owner, string text) : base(owner, text)
        {
            ItemType = TrayMenuItemType.CheckBox;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        /// <param name="visible">item visibility</param>
        internal TrayMenuCheckboxItem(TrayMenu owner, string text, bool visible) : base(owner, text, visible)
        {
            ItemType = TrayMenuItemType.CheckBox;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Checked State
        /// </summary>
        public bool Checked
        {
            get
            {
                return _checked;
            }
            set
            {
                if (_checked != value)
                {
                    _checked = value;
                    Owner.OnItemCheckedChanged(this);
                }
            }
        }

        #endregion

        #region Overrides

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

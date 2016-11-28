using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents a tray menu auto close checkbox item
    /// </summary>
    [ItemType(TrayMenuItemType.AutoClose)]
    public class TrayMenuAutoCloseItem : TrayMenuItem
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        internal TrayMenuAutoCloseItem(TrayMenu owner, string text) : base(owner, text)
        {
            ItemType = TrayMenuItemType.AutoClose;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        /// <param name="visible">item visibility</param>
        internal TrayMenuAutoCloseItem(TrayMenu owner, string text, bool visible) : base(owner, text, visible)
        {
            ItemType = TrayMenuItemType.AutoClose;
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

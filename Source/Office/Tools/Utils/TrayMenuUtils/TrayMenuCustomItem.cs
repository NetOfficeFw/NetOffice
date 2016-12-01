using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents a tray menu custom item
    /// </summary>
    [ItemType(TrayMenuItemType.Custom)]
    public class TrayMenuCustomItem : TrayMenuItem
    {
        #region Fields

        private object _control;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        /// <param name="visible">item visibility</param>
        /// <param name="control">custom control</param>
        internal TrayMenuCustomItem(TrayMenu owner, string text, bool visible, object control) : base(owner, text, visible)
        {
            if (null == control)
                throw new ArgumentNullException("control");
            _control = control;
            ItemType = TrayMenuItemType.Custom;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Custom Control
        /// </summary>
        public object Control
        {
            get
            {
                return _control;
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

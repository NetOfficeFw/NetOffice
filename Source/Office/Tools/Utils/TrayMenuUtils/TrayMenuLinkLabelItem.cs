using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents a tray menu link label item
    /// </summary>
    [ItemType(TrayMenuItemType.LinkLabel)]
    public class TrayMenuLinkLabelItem : TrayMenuItem
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        internal TrayMenuLinkLabelItem(TrayMenu owner, string text) : base(owner, text)
        {
            ItemType = TrayMenuItemType.LinkLabel;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        /// <param name="visible">item visibility</param>
        internal TrayMenuLinkLabelItem(TrayMenu owner, string text, bool visible) : base(owner, text, visible)
        {
            ItemType = TrayMenuItemType.LinkLabel;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Shown Text Alignment
        /// </summary>
        public override ContentAlignment TextAlign
        {
            get
            {
                return base.TextAlign;
            }

            set
            {
                base.TextAlign = value;
            }
        }

        /// <summary>
        /// Shown Image Alignment
        /// </summary>
        public override ContentAlignment ImageAlign
        {
            get
            {
                return base.ImageAlign;
            }

            set
            {
                base.ImageAlign = value;
            }
        }

        /// <summary>
        /// Padding Space
        /// </summary>
        public override Padding Padding
        {
            get
            {
                return base.Padding;
            }

            set
            {
                base.Padding = value;
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
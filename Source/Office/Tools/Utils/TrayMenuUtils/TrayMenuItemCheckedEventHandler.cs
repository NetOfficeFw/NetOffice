using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Tray menu item related checked arguments
    /// </summary>
    public class TrayMenuItemCheckedEventArgs : TrayMenuItemsEventArgs
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="item">target item</param>
        /// <param name="check">item checked state</param>
        public TrayMenuItemCheckedEventArgs(TrayMenuItem item, bool check) : base(item)
        {
            Checked = check;
        }

        /// <summary>
        /// Indicates the item is checked
        /// </summary>
        public bool Checked { get; private set; }
    }

    /// <summary>
    /// Tray menu item check event handler
    /// </summary>
    /// <param name="sender">sender instance</param>
    /// <param name="args">click event arguments</param>
    public delegate void TrayMenuItemCheckedEventHandler(object sender, TrayMenuItemCheckedEventArgs args);
}

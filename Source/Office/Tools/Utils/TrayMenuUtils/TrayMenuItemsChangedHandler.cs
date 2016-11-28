using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Tray menu item related arguments
    /// </summary>
    public class TrayMenuItemsEventArgs : EventArgs
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="item">tray menu item</param>
        public TrayMenuItemsEventArgs(TrayMenuItem item)
        {
            Item = item;
        }

        /// <summary>
        /// Tray menu item
        /// </summary>
        public TrayMenuItem Item { get; private set; }
    }

    /// <summary>
    /// Item related changed event handler
    /// </summary>
    /// <param name="sender">sender instance</param>
    /// <param name="args">item arguments</param>
    public delegate void TrayMenuItemsChangedHandler(object sender, TrayMenuItemsEventArgs args);
}

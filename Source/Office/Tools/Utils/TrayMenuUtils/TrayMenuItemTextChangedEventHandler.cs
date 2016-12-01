using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Item Text Change Event Handler
    /// </summary>
    /// <param name="sender">sender instance</param>
    /// <param name="args">changed arguments</param>
    public delegate void TrayMenuItemTextChangedEventHandler(object sender, TrayMenuItemTextChangedEventArgs args);

    /// <summary>
    /// Item Text Change Event Arguments
    /// </summary>
    public class TrayMenuItemTextChangedEventArgs : EventArgs
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="item"></param>
        public TrayMenuItemTextChangedEventArgs(TrayMenuItem item)
        {
            Item = item;
        }

        /// <summary>
        /// Text changed item
        /// </summary>
        public TrayMenuItem Item { get; private set; }
    }
}

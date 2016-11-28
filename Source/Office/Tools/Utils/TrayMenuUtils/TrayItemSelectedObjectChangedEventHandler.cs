using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Tray menu item related arguments
    /// </summary>
    public class TrayMenuItemSelectedObjectChangedEventArgs : EventArgs
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="item">tray menu item</param>
        /// <param name="selectedObject">current selected item</param>
        /// <param name="selectedIndex">current selected index</param>
        public TrayMenuItemSelectedObjectChangedEventArgs(TrayMenuItem item, object selectedObject, int selectedIndex)
        {
            Item = item;
            SelectedObject = selectedObject;
            SelectedIndex = selectedIndex;            
        }

        /// <summary>
        /// Tray menu item
        /// </summary>
        public TrayMenuItem Item { get; private set; }

        /// <summary>
        /// Current Selected Object
        /// </summary>
        public object SelectedObject { get; private set; }

        /// <summary>
        /// Current Selected Index
        /// </summary>
        public int SelectedIndex { get; private set; }
    }

    /// <summary>
    /// Tray menu item selected object changed event handler
    /// </summary>
    /// <param name="instance">sender instance</param>
    /// <param name="args">changed arguments</param>
    public delegate void TrayMenuItemSelectedObjectChangedEventHandler(object instance, TrayMenuItemSelectedObjectChangedEventArgs args);     
}

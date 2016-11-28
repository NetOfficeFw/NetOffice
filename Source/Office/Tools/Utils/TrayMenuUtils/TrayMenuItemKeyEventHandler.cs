using System;
using System.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Provides arguments for the TrayMenuItemKey handler
    /// </summary>
    public class TrayMenuItemKeyEventArgs : EventArgs
    {
        /// <summary>
        ///  Creates a new instance of the class
        /// </summary>
        /// <param name="item">event item</param>
        /// <param name="keyData">ToolsKeys representing the key that was pressed</param>
        /// <param name="alt">indicating whether the ALT key was pressed</param>
        /// <param name="control">indicating whether the CTRL key was pressed</param>
        /// <param name="handled">indicating whether the event was handled</param>
        /// <param name="keyCode">keyboard code for a KeyDown or KeyUp event</param>
        /// <param name="keyValue">integer representation of the KeyCode property</param>
        /// <param name="modifiers">modifier flags for a KeyDown or KeyUp event</param>
        /// <param name="shift">indicating whether the SHIFT key was pressed</param>
        /// <param name="suppressKeyPress">indicating whether the key event should be passed on to the underlying control</param>
        [TargetedPatchingOptOut("Performance critical to inline this type of method across NGen image boundaries")]
        public TrayMenuItemKeyEventArgs(TrayMenuItem item, ToolsKeys keyData, bool alt, bool control, bool handled, ToolsKeys keyCode, int keyValue, ToolsKeys modifiers, bool shift, bool suppressKeyPress )
        {
            Item = item;
            Alt = alt;
            Control = control;
            Handled = handled;
            KeyCode = keyCode;
            KeyData = keyData;
            KeyValue = keyValue;
            Modifiers = modifiers;
            Shift = shift;
            SuppressKeyPress = suppressKeyPress;
        }

        /// <summary>
        /// Event Item
        /// </summary>
        public TrayMenuItem Item { get; private set; }

        /// <summary>
        ///  Gets a value indicating whether the ALT key was pressed.
        /// </summary>
        public virtual bool Alt { get; }

        /// <summary>
        ///  Gets a value indicating whether the CTRL key was pressed.
        /// </summary>
        public bool Control { get; }

        /// <summary>
        /// Gets or sets a value indicating whether the event was handled.
        /// </summary>
        public bool Handled { get; set; }

        /// <summary>
        /// Gets the keyboard code for a KeyDown or KeyUp event
        /// </summary>
        public ToolsKeys KeyCode { get; }

        /// <summary>
        /// Gets the key data for a KeyDown or KeyUp event.
        /// </summary>
        public ToolsKeys KeyData { get; }

        /// <summary>
        /// The integer representation of the KeyCode property.
        /// </summary>
        public int KeyValue { get; }

        /// <summary>
        /// Gets the modifier flags for a KeyDown or KeyUp event.
        /// </summary>
        public ToolsKeys Modifiers { get; }

        /// <summary>
        ///  Gets a value indicating whether the SHIFT key was pressed.
        /// </summary>
        public virtual bool Shift { get; }

        /// <summary>
        /// Gets or sets a value indicating whether the key event should be passed on to the underlying control.
        /// </summary>
        public bool SuppressKeyPress { get; set; }
    }

    /// <summary>
    /// Key Down/Up Event Handler
    /// </summary>
    /// <param name="sender">sender instance</param>
    /// <param name="args">key arguments</param>
    public delegate void TrayMenuItemKeyEventHandler(object sender, TrayMenuItemKeyEventArgs args);    
}
